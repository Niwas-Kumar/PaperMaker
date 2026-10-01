#!/usr/bin/env python3
"""Paper Maker core engine — parse DOCX papers and build trimmed/reordered copies.

Shared by the GUI (``paper_maker.py``) and the CLI (``extract_questions.py``).

Strategy: copy the whole source package (keeping every style, image, equation,
header and footer), then rebuild the document body so it contains only the
selected questions, in the chosen order, inside a single continuous two-column
section that reuses the original header (on every page) and footer.
"""

from __future__ import annotations

import copy
import re
from collections import OrderedDict
from collections.abc import Iterable, Iterator
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURABLE PATTERNS — edit these to match your document conventions
# ═══════════════════════════════════════════════════════════════════════════════

SUBJECT_KEYWORDS: tuple[str, ...] = (
    "PHYSICS",
    "CHEMISTRY",
    "BIOLOGY",
    "BOTANY",
    "ZOOLOGY",
    "MATHEMATICS",
    "ENGLISH",
    "HINDI",
)

# Group 1 must capture the question number. The lookahead stops a stem line such
# as "2.5 m/s …" from being mistaken for question 2.
QUESTION_RE = re.compile(r"^\s*(\d+)\.(?!\d)")

# A short line that opens an answer-key / solutions part. Everything from here on
# repeats the question numbers, so segmentation stops at it.
ANSWER_MARKER_RE = re.compile(
    r"^\s*(?:answer\s*keys?|answers?|solutions?|"
    r"hints?(?:\s*(?:&|and)\s*solutions?)?|"
    r"answer\s*keys?\s*(?:&|and)\s*solutions?)\s*:?\s*$",
    re.IGNORECASE,
)

# A solution block opens with the key-line form "N.(k)" — the number AND the
# answer — so numbered steps inside an explanation can never start a new block.
SOLUTION_START_RE = re.compile(r"^\s*(\d+)\.\s*\(\s*[^)\s]{1,6}\s*\)")

SegmentKey = tuple[str, int]  # (subject, question number)


def _implicit_subject(path: str | Path) -> str:
    """Subject label for a paper that has no subject headings of its own.

    Single-subject papers (e.g. ``…_Botany_…Question.docx``) name the subject in
    the file name; otherwise fall back to a neutral label.
    """
    stem = Path(path).stem.upper()
    for kw in SUBJECT_KEYWORDS:
        if kw in stem:
            return kw.title()
    return "Questions"


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT HELPERS
# ═══════════════════════════════════════════════════════════════════════════════


def iter_block_items(document: DocumentObject) -> Iterator[Paragraph | Table]:
    """Yield body-level Paragraphs and Tables in document order."""
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def _is_subject_heading(paragraph: Paragraph) -> str | None:
    """Return the normalised subject name if *paragraph* is a subject heading."""
    text = paragraph.text.strip()
    if not text:
        return None

    upper = text.upper()
    for kw in SUBJECT_KEYWORDS:
        if upper == kw:
            return kw.title()

    m = re.match(r"^Subject:\s*(.*)", text, re.IGNORECASE)
    if m:
        return m.group(1).strip().title()

    # Style-based fallback is expensive (resolves the styles part), so only
    # consult it for text short enough to plausibly be a heading.
    if len(text) < 60:
        style_name = paragraph.style.name if paragraph.style else ""
        if style_name.startswith("Heading"):
            is_bold = bool(paragraph.runs) and all(
                r.bold for r in paragraph.runs if r.text.strip()
            )
            if is_bold:
                return text.title()

    return None


# ═══════════════════════════════════════════════════════════════════════════════
# QUESTION MODEL + PARSING
# ═══════════════════════════════════════════════════════════════════════════════


class QuestionInfo:
    """Metadata about a single detected question."""

    __slots__ = ("number", "subject", "preview")

    def __init__(self, number: int, subject: str, preview: str) -> None:
        self.number = number
        self.subject = subject
        self.preview = preview

    def label(self) -> str:
        return f"Q{self.number}. {self.preview}"

    def __repr__(self) -> str:
        return f"QuestionInfo(number={self.number}, subject={self.subject!r})"


class ParsedPaper:
    """What the GUI/CLI learn about a source file before building anything."""

    __slots__ = ("subjects", "has_answer_key")

    def __init__(self, subjects, has_answer_key: bool) -> None:
        self.subjects: OrderedDict[str, list[QuestionInfo]] = subjects
        self.has_answer_key = has_answer_key


def parse_paper(path: str | Path) -> ParsedPaper:
    """Parse a .docx into ``ParsedPaper(subjects, has_answer_key)``.

    Uses the same segmenter as the builder, so the questions offered in the GUI
    are exactly the blocks that end up in the paper — for both the multi-subject
    (heading-per-subject) layout and single-subject papers with no headings.
    ``has_answer_key`` is true when the source carries an ANSWER KEY part.
    """
    doc = Document(str(path))
    parsed = _segment_questions(doc, _implicit_subject(path))

    subjects: OrderedDict[str, list[QuestionInfo]] = OrderedDict()
    for (subject, number), elements in parsed.segments.items():
        text = ""
        for el in elements:
            if el.tag == qn("w:p"):
                text = Paragraph(el, doc).text.strip()
                break
        m = QUESTION_RE.match(text)
        raw = text[m.end():].strip() if m else text
        preview = (raw[:77] + "…") if len(raw) > 80 else raw
        subjects.setdefault(subject, []).append(
            QuestionInfo(number=number, subject=subject, preview=preview)
        )
    return ParsedPaper(subjects, bool(parsed.answers))


def parse_document(path: str | Path) -> "OrderedDict[str, list[QuestionInfo]]":
    """Parse a .docx and return ``{subject: [QuestionInfo, …]}``."""
    return parse_paper(path).subjects


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACT & REBUILD ENGINE
# ═══════════════════════════════════════════════════════════════════════════════


class _Parsed:
    """Everything the segmenter harvests from one source document."""

    __slots__ = ("body", "sectPr", "preamble", "heading_templates", "segments",
                 "answer_head", "answers", "solution_head", "solutions")

    def __init__(self, body, sectPr, preamble, heading_templates, segments,
                 answer_head, answers, solution_head, solutions) -> None:
        self.body = body
        self.sectPr = sectPr
        self.preamble = preamble
        self.heading_templates = heading_templates
        self.segments = segments
        self.answer_head = answer_head
        self.answers = answers
        self.solution_head = solution_head  # elements between key and solutions
        self.solutions = solutions          # {number: [elements]}


def _carries_section_break(paragraph_el) -> bool:
    pPr = paragraph_el.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:sectPr")) is not None


def _segment_questions(
    doc: DocumentObject,
    implicit_subject: str = "Questions",
) -> _Parsed:
    """Split the document body into the preamble and per-question runs.

    Returns ``(body, sectPr, preamble, heading_templates, segments)``:

    * *preamble* — the body-level elements before the first question block. For
      the multi-subject layout that is the cover (title banner, syllabus table,
      instructions); for a single-subject paper it is the syllabus line and its
      section break. Either way it is reused verbatim so the output mirrors the
      source's own first page.
    * *heading_templates* — ``{subject: heading_paragraph_element}`` for the
      first occurrence of each subject banner (empty for papers that have no
      subject headings — those get one implicit subject, *implicit_subject*).
    * *segments* — ``{(subject, number): [elements]}`` for each question.

    Question segmentation stops at an answer-key / solutions part (a marker line
    such as "ANSWER KEY", or the question numbering restarting), so those never
    get appended onto the questions. When a marker is found, the key itself is
    harvested too: *answer_head* is the marker paragraph through the section
    break that precedes the first answer line, and *answers* maps each question
    number to its one-line answer paragraph (``136.(1)``); the key ends at the
    first answer-region paragraph carrying a section break or when its numbering
    restarts (the solutions part). The terminal body-level ``w:sectPr`` is
    returned separately.
    """
    body = doc.element.body
    sectPr = body.find(qn("w:sectPr"))
    children = list(body.iterchildren())

    # Decide the layout up front: a paper with subject headings keeps every
    # numbered line before its first heading (the cover's instructions) in the
    # preamble; only a paper with no headings at all is single-subject.
    heading_of: dict[int, str | None] = {}
    for child in children:
        if isinstance(child, CT_P):
            heading_of[id(child)] = _is_subject_heading(Paragraph(child, doc))
    headerless = not any(heading_of.values())

    preamble: list = []
    heading_templates: "OrderedDict[str, object]" = OrderedDict()
    segments: OrderedDict[SegmentKey, list] = OrderedDict()
    answer_head: list = []
    answers: "OrderedDict[int, object]" = OrderedDict()
    solution_head: list = []
    solutions: "OrderedDict[int, list]" = OrderedDict()
    current_subject: str | None = None
    found_subject = False
    current_key: SegmentKey | None = None
    last_number: int | None = None
    last_answer: int | None = None
    last_solution: int | None = None
    current_solution: int | None = None
    phase = "questions"  # → "answer_head" → "answers" → "solutions" → "done"

    for child in children:
        if child is sectPr or phase == "done":
            continue
        is_para = isinstance(child, CT_P)
        text = Paragraph(child, doc).text.strip() if is_para else ""

        if phase == "answer_head":
            if is_para and QUESTION_RE.match(text):
                phase = "answers"  # first answer line — fall through
            else:
                answer_head.append(child)
                continue

        if phase == "answers":
            m = QUESTION_RE.match(text) if is_para else None
            if m and last_answer is not None and int(m.group(1)) <= last_answer:
                phase = "solutions"  # numbering restarted: this IS a solution
            else:
                if m:
                    answers.setdefault(int(m.group(1)), child)
                    last_answer = int(m.group(1))
                if is_para and _carries_section_break(child):
                    phase = "solutions"  # the key's section ends here
                continue

        if phase == "solutions":
            # Solutions follow the key (the user's papers always carry them):
            # one block per question, opened by a "N.(k)" line.
            m = SOLUTION_START_RE.match(text) if is_para else None
            if m and (last_solution is None or int(m.group(1)) > last_solution):
                current_solution = int(m.group(1))
                last_solution = current_solution
                solutions.setdefault(current_solution, [])
            if current_solution is None:
                solution_head.append(child)  # blank lines / break before them
            else:
                solutions[current_solution].append(child)
            continue

        if is_para:
            subj = heading_of.get(id(child))
            if subj:
                current_subject = subj
                found_subject = True
                current_key = None
                heading_templates.setdefault(subj, child)  # keep the banner
                continue  # drop the heading paragraph itself here

            if found_subject and ANSWER_MARKER_RE.match(text):
                answer_head.append(child)
                phase = "answer_head"
                continue

            m = QUESTION_RE.match(text)
            if m and (found_subject or headerless):
                number = int(m.group(1))
                if (found_subject and last_number is not None
                        and number <= last_number):
                    phase = "done"  # numbering restarted, no marker seen
                    continue
                if not found_subject:
                    # No subject headings in this paper: one implicit subject,
                    # and everything seen so far is the preamble.
                    current_subject = implicit_subject
                    found_subject = True
                current_key = (current_subject, number)
                last_number = number
                segments.setdefault(current_key, [])

        if not found_subject:
            preamble.append(child)
            continue

        if current_key is not None:
            segments[current_key].append(child)

    return _Parsed(body, sectPr, preamble, heading_templates, segments,
                   answer_head, answers, solution_head, solutions)


def _strip_section_break(paragraph_el) -> bool:
    """Remove an embedded ``w:sectPr`` (a mid-document section break).

    Returns ``True`` when the paragraph existed only to carry that break (no runs
    of its own) and should therefore be dropped entirely.
    """
    pPr = paragraph_el.find(qn("w:pPr"))
    had_break = False
    if pPr is not None:
        sect = pPr.find(qn("w:sectPr"))
        if sect is not None:
            pPr.remove(sect)
            had_break = True
    has_content = paragraph_el.find(qn("w:r")) is not None
    return had_break and not has_content


def _ensure_ref(sectPr, tag: str, type_val: str, rid: str) -> None:
    """Add a header/footer reference to *sectPr* if one of that type is absent.

    References must precede all other section properties in the schema, so the
    element is inserted after any existing references and before everything else.
    """
    qtag = qn(tag)
    for existing in sectPr.findall(qtag):
        if existing.get(qn("w:type")) == type_val:
            return

    el = OxmlElement(tag)
    el.set(qn("w:type"), type_val)
    el.set(qn("r:id"), rid)

    insert_at = 0
    ref_tags = (qn("w:headerReference"), qn("w:footerReference"))
    for i, node in enumerate(sectPr):
        if node.tag in ref_tags:
            insert_at = i + 1
        else:
            break
    sectPr.insert(insert_at, el)


def _discover_layout_refs(
    doc: DocumentObject,
) -> tuple[list[str], list[tuple[str, str]]]:
    """Collect header/footer relationship ids used anywhere in the source.

    They live on the source's section breaks, so scan every ``w:sectPr``.
    Returns ``(header_rids, footer_refs)`` where *header_rids* is the distinct
    header relationship ids in document order and *footer_refs* is a normalised
    ``[(type, rid), …]`` covering both odd (default) and even pages.
    """
    body = doc.element.body
    header_rids: list[str] = []
    footers: dict[str, str] = {}

    for sectPr in body.iter(qn("w:sectPr")):
        for ref in sectPr.findall(qn("w:headerReference")):
            rid = ref.get(qn("r:id"))
            if rid and rid not in header_rids:
                header_rids.append(rid)
        for ref in sectPr.findall(qn("w:footerReference")):
            rid = ref.get(qn("r:id"))
            type_val = ref.get(qn("w:type")) or "default"
            if rid:
                footers.setdefault(type_val, rid)

    default_f = footers.get("default") or footers.get("first") or footers.get("even")
    even_f = footers.get("even") or default_f
    footer_refs: list[tuple[str, str]] = []
    if default_f:
        footer_refs.append(("default", default_f))
    if even_f:
        footer_refs.append(("even", even_f))
    return header_rids, footer_refs


def _set_cols(sectPr, num: int) -> None:
    """Set the column count of *sectPr* (``num`` = 1 for a full-width section)."""
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sectPr.append(cols)
    cols.set(qn("w:num"), str(max(1, num)))
    cols.set(qn("w:space"), "709")
    if num > 1:
        cols.set(qn("w:sep"), "1")
    elif qn("w:sep") in cols.attrib:
        del cols.attrib[qn("w:sep")]


def _set_type(sectPr, val: str) -> None:
    """Set the section-break ``w:type`` (e.g. ``continuous``), kept after refs."""
    for existing in sectPr.findall(qn("w:type")):
        sectPr.remove(existing)
    el = OxmlElement("w:type")
    el.set(qn("w:val"), val)
    insert_at = 0
    ref_tags = (qn("w:headerReference"), qn("w:footerReference"))
    for i, node in enumerate(sectPr):
        if node.tag in ref_tags:
            insert_at = i + 1
        else:
            break
    sectPr.insert(insert_at, el)


def _attach_footer_refs(
    sectPr,
    footer_refs: list[tuple[str, str]],
) -> None:
    """Attach the page-number footer(s) to *sectPr*.

    No header reference is added: the source has no running header on its content
    pages (the institute logo lives only in the cover's title box), so rebuilt
    sections stay header-free to match the original exactly.
    """
    for type_val, rid in footer_refs:
        _ensure_ref(sectPr, "w:footerReference", type_val, rid)


def _fresh_sectpr(base_sectPr) -> object:
    """Deep-copy the body *base_sectPr* (page size/margins) stripped of layout.

    Header/footer refs, the break type and ``titlePg`` are removed so callers can
    set exactly the layout they need; the page geometry is preserved.
    """
    sect = copy.deepcopy(base_sectPr)
    for tag in ("w:headerReference", "w:footerReference", "w:type", "w:titlePg"):
        for el in sect.findall(qn(tag)):
            sect.remove(el)
    return sect


def _section_break_para(
    base_sectPr,
    *,
    cols: int,
    footer_refs: list[tuple[str, str]],
) -> object:
    """Build an empty paragraph that carries a continuous section break.

    The break ends the preceding run of content as a *cols*-column section that
    keeps the page-number footer (and no running header, like the source).
    """
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    p.append(pPr)
    sect = _fresh_sectpr(base_sectPr)
    _set_cols(sect, cols)
    _set_type(sect, "continuous")
    _attach_footer_refs(sect, footer_refs)
    pPr.append(sect)
    return p


def _heading_para(
    template,
    subject: str,
    base_sectPr,
    footer_refs: list[tuple[str, str]],
) -> object:
    """Build a full-width subject banner that ends a single-column section.

    Reuses the original heading paragraph (grey bar, font, shading) when one is
    available; otherwise synthesises a bold uppercase heading.
    """
    if template is not None:
        p = copy.deepcopy(template)
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)
        old = pPr.find(qn("w:sectPr"))
        if old is not None:
            pPr.remove(old)
    else:
        p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        p.append(pPr)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center")
        pPr.append(jc)
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        b = OxmlElement("w:b")
        rpr.append(b)
        r.append(rpr)
        t = OxmlElement("w:t")
        t.text = subject.upper()
        r.append(t)
        p.append(r)

    sect = _fresh_sectpr(base_sectPr)
    _set_cols(sect, 1)
    _set_type(sect, "continuous")
    _attach_footer_refs(sect, footer_refs)
    pPr.append(sect)
    return p


def _renumber_segment(elements: list, new_number: int) -> None:
    """Rewrite the leading ``N.`` of a question block to *new_number*.

    The number opens the first paragraph but may be split across runs (one
    source has ``'4'`` + ``'6.'``), so leading runs are joined until they form
    ``N.``; the new number goes into the first of them and the rest are emptied.
    Options such as ``(1)`` are untouched because only the first paragraph is
    examined.
    """
    for el in elements:
        if el.tag != qn("w:p"):
            continue
        texts = el.findall(qn("w:r") + "/" + qn("w:t"))
        joined = ""
        for i, t in enumerate(texts):
            joined += t.text or ""
            m = QUESTION_RE.match(joined)
            if m:
                rest = joined[m.end():]
                texts[0].text = f"{new_number}."
                for other in texts[1:i]:
                    other.text = ""
                if i > 0:
                    texts[i].text = rest
                else:
                    texts[0].text += rest
                return
            if len(joined) > 16:
                break  # not a number prefix — leave the block as-is
        return  # first paragraph had no number — leave the block as-is


def _join_names(names: list[str]) -> str:
    """Join subject names as ``A``, ``A and B`` or ``A, B and C``."""
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    return f"{', '.join(names[:-1])} and {names[-1]}"


def _trim_syllabus(tbl, selected_lower: set[str]) -> None:
    """Drop syllabus rows whose subject was not selected."""
    for tr in list(tbl.findall(qn("w:tr"))):
        tc = tr.find(qn("w:tc"))
        label = ""
        if tc is not None:
            label = "".join(t.text or "" for t in tc.iter(qn("w:t"))).lower()
        row_subj = None
        for kw in SUBJECT_KEYWORDS:
            if kw.lower() in label:
                row_subj = kw.lower()
                break
        if row_subj is not None and row_subj not in selected_lower:
            tbl.remove(tr)


def _customize_preamble(
    preamble: list,
    subjects_order: list[str],
    total: int,
) -> None:
    """Rewrite the cover page's instruction counts and syllabus for this paper.

    The question total, the maximum mark (``4 × total``) and the subject list are
    all made to reflect the questions the teacher actually selected.
    """
    marks = total * 4
    phrase = _join_names([s for s in subjects_order])
    selected_lower = {s.lower() for s in subjects_order}

    for el in preamble:
        if el.tag == qn("w:tbl"):
            _trim_syllabus(el, selected_lower)
            continue
        if el.tag != qn("w:p"):
            continue
        for t in el.findall(qn("w:r") + "/" + qn("w:t")):
            if not t.text:
                continue
            if "multiple-choice question" in t.text:
                t.text = re.sub(r"\d+\s+multiple-choice question",
                                f"{total} multiple-choice question", t.text)
            if phrase and re.search(r"from\s+Physics,", t.text):
                t.text = re.sub(r"from\s+.*?\.\s*All questions",
                                f"from {phrase}. All questions", t.text)
            elif t.text.strip() == "720":
                t.text = t.text.replace("720", str(marks))


def _build(
    source: Path,
    output: Path,
    pick: "callable[[OrderedDict[SegmentKey, list]], list]",
) -> None:
    """Copy *source* → *output* and rebuild its body from picked segments.

    *pick* receives the ``{(subject, number): [elements]}`` map and returns the
    ordered list of element-lists to keep.

    The source is opened directly (never a fresh copy) and written straight to
    *output*. Reading the freshly-copied file would force the on-access malware
    scan of its hundreds of embedded OLE objects — ~8s per file — whereas the
    original is already trusted, so this keeps generation to ~1-2s.

    The rebuilt body keeps the real cover page (title banner, syllabus and
    instructions — with the counts/subjects rewritten for this paper), inserts a
    full-width grey heading before each subject's block, and renumbers the kept
    questions ``1, 2, 3, …`` in the teacher's chosen order. Each subject banner
    sits in its own single-column section and each question block in a two-column
    section, exactly like the source, so the banner header shows on every page.
    """
    doc = Document(str(source))
    _, footer_refs = _discover_layout_refs(doc)
    parsed = _segment_questions(doc, _implicit_subject(source))
    body, sectPr, preamble = parsed.body, parsed.sectPr, parsed.preamble
    heading_templates, segments = parsed.heading_templates, parsed.segments
    # Papers without subject headings (single-subject tests) get no cover
    # rewrite and no heading bars: their own first page is mirrored as-is.
    has_headings = bool(heading_templates)

    picked = [(subj, els) for subj, els in pick(segments) if els]

    subjects_order: list[str] = []
    for subj, _els in picked:
        if subj not in subjects_order:
            subjects_order.append(subj)
    total = len(picked)

    for child in list(body.iterchildren()):
        if child is not sectPr:
            body.remove(child)

    # Cover page (first single-column section with its own footer / title page),
    # reused verbatim so the title box, logo and spacing match the source exactly.
    if preamble:
        if has_headings:
            _customize_preamble(preamble, subjects_order, total)
        for el in preamble:
            sectPr.addprevious(el)

    prev_subject: str | None = None
    number = 0
    for subj, element_list in picked:
        if subj != prev_subject and has_headings:
            if prev_subject is not None:
                # Close the previous subject's two-column question section.
                sectPr.addprevious(_section_break_para(
                    sectPr, cols=2, footer_refs=footer_refs))
            sectPr.addprevious(_heading_para(
                heading_templates.get(subj), subj, sectPr, footer_refs))
        prev_subject = subj

        number += 1
        _renumber_segment(element_list, number)
        for el in element_list:
            if el.tag == qn("w:p") and _strip_section_break(el):
                continue
            sectPr.addprevious(el)

    # The terminal body section closes the last subject's two-column block.
    _set_cols(sectPr, 2)
    _set_type(sectPr, "continuous")
    title_pg = sectPr.find(qn("w:titlePg"))
    if title_pg is not None:
        sectPr.remove(title_pg)
    _attach_footer_refs(sectPr, footer_refs)
    return _save_with_fallback(doc, output)


def _save_with_fallback(doc: DocumentObject, output: Path) -> Path:
    """Save *doc* to *output*, or to ``name (2).docx`` … if *output* is locked.

    A paper the user still has open in Word cannot be overwritten (Windows
    raises ``PermissionError``), so rather than failing the whole batch the
    file is written under the next free numbered name. Returns the path used.
    """
    try:
        doc.save(str(output))
        return output
    except PermissionError:
        pass
    for n in range(2, 50):
        candidate = output.with_name(f"{output.stem} ({n}){output.suffix}")
        if candidate.exists():
            continue
        try:
            doc.save(str(candidate))
            return candidate
        except PermissionError:
            continue
    raise PermissionError(
        f"Cannot write {output.name}: it is open in another program (close it "
        "in Word and generate again).")


def generate_paper(
    source_path: str | Path,
    selected_questions: list[QuestionInfo],
    output_path: str | Path,
    progress_callback=None,
) -> None:
    """Build one paper from *selected_questions* (subject-aware, custom order)."""
    if not selected_questions:
        raise ValueError("No questions selected.")

    source = Path(source_path)
    output = Path(output_path)
    total = len(selected_questions)

    def pick(segments):
        kept = []
        for idx, qi in enumerate(selected_questions, 1):
            kept.append((qi.subject, segments.get((qi.subject, qi.number), [])))
            if progress_callback:
                progress_callback(idx, total)
        return kept

    return _build(source, output, pick)


def generate_answer_key(
    source_path: str | Path,
    selected_questions: list[QuestionInfo],
    output_path: str | Path,
) -> Path:
    """Build a separate answer-key paper for *selected_questions*.

    The source's own ANSWER KEY part supplies the answer lines; the output keeps
    the source's first page and the key's own heading/section layout, and the
    answers are renumbered to match the question paper built from the same
    selection (``1, 2, 3, …`` in the chosen order). Raises ``ValueError`` when
    the source has no answer key. Returns the path actually written.
    """
    if not selected_questions:
        raise ValueError("No questions selected.")
    source = Path(source_path)
    output = Path(output_path)

    doc = Document(str(source))
    _, footer_refs = _discover_layout_refs(doc)
    parsed = _segment_questions(doc, _implicit_subject(source))
    if not parsed.answers:
        raise ValueError("No answer key was detected in this paper.")

    body, sectPr = parsed.body, parsed.sectPr
    for child in list(body.iterchildren()):
        if child is not sectPr:
            body.remove(child)
    for el in parsed.preamble:
        sectPr.addprevious(el)
    for el in parsed.answer_head:
        sectPr.addprevious(el)

    # The paper's numbering: a selection missing from the paper consumes none.
    numbered: list[tuple[int, QuestionInfo]] = []
    for qi in selected_questions:
        if parsed.segments.get((qi.subject, qi.number)):
            numbered.append((len(numbered) + 1, qi))

    for number, qi in numbered:
        answer = parsed.answers.get(qi.number)
        if answer is None:
            continue  # question exists but the key has no line for it
        answer = copy.deepcopy(answer)
        _strip_section_break(answer)
        _renumber_segment([answer], number)
        sectPr.addprevious(answer)

    # Solutions, when the source has them: close the key's two-column section
    # the way the source does, keep the source's own gap, then the selected
    # blocks under the same new numbers.
    if parsed.solutions:
        sectPr.addprevious(_section_break_para(
            sectPr, cols=2, footer_refs=footer_refs))
        for el in parsed.solution_head:
            sectPr.addprevious(el)
        for number, qi in numbered:
            block = parsed.solutions.get(qi.number)
            if not block:
                continue
            _renumber_segment(block, number)
            for el in block:
                if el.tag == qn("w:p") and _strip_section_break(el):
                    continue
                sectPr.addprevious(el)

    _set_cols(sectPr, 2)
    _set_type(sectPr, "continuous")
    title_pg = sectPr.find(qn("w:titlePg"))
    if title_pg is not None:
        sectPr.remove(title_pg)
    _attach_footer_refs(sectPr, footer_refs)
    return _save_with_fallback(doc, output)


def extract_questions_by_number(
    source_path: str | Path,
    numbers: Iterable[int],
    output_path: str | Path,
) -> tuple[set[int], list[int]]:
    """Keep only *numbers* (ignoring subject) in the given order.

    Returns ``(detected_numbers, missing_numbers)``.
    """
    source = Path(source_path)
    output = Path(output_path)
    requested = list(dict.fromkeys(int(n) for n in numbers))  # de-dup, keep order
    if not requested:
        raise ValueError("At least one question number must be requested.")

    detected: set[int] = set()

    def pick(segments):
        by_number: dict[int, tuple[str, list]] = {}
        for (subject, num), elements in segments.items():
            detected.add(num)
            by_number.setdefault(num, (subject, elements))  # first wins
        return [by_number[n] for n in requested if n in by_number]

    _build(source, output, pick)
    missing = [n for n in requested if n not in detected]
    return detected, missing


# ═══════════════════════════════════════════════════════════════════════════════
# BATCH — one paper per source file, generated in parallel
# ═══════════════════════════════════════════════════════════════════════════════


MODE_QUESTIONS = "questions"   # question paper only
MODE_ANSWERS = "answers"       # answer key only
MODE_BOTH = "both"             # question paper + answer key


class BatchJob:
    """A single file's generation request within a batch.

    *mode* chooses which papers to write (see ``MODE_*``); the answer key goes to
    *answer_output*, defaulting to ``<output stem>_answer_key.docx``. After the
    run, *saved* lists the paths actually written (they can differ from the
    requested ones when a target was open in Word).
    """

    __slots__ = ("source", "questions", "output", "mode", "answer_output", "saved")

    def __init__(
        self,
        source: str | Path,
        questions: list[QuestionInfo],
        output: str | Path,
        mode: str = MODE_QUESTIONS,
        answer_output: str | Path | None = None,
    ) -> None:
        self.source = Path(source)
        self.questions = questions
        self.output = Path(output)
        self.mode = mode
        self.answer_output = Path(answer_output) if answer_output else None
        self.saved: list[Path] = []


def generate_papers_batch(
    jobs: list[BatchJob],
    max_workers: int | None = None,
    progress_callback=None,
) -> list[tuple[BatchJob, Exception | None]]:
    """Generate every job's paper concurrently.

    ``progress_callback(done, total, job, error)`` is invoked as each file
    finishes (``error`` is ``None`` on success). Returns ``[(job, error), …]``.
    File copy and save are I/O-bound, so threads overlap the slow parts even
    under the GIL. Returns results in completion order.
    """
    if not jobs:
        raise ValueError("No files to generate.")

    if max_workers is None:
        max_workers = min(len(jobs), 8)
    max_workers = max(1, min(max_workers, len(jobs)))

    results: list[tuple[BatchJob, Exception | None]] = []
    total = len(jobs)

    def run(job: BatchJob) -> None:
        job.saved = []
        if job.source.suffix.lower() == ".pdf":
            from pdf_core import generate_pdf_answer_key, generate_pdf_paper
            make_paper, make_key = generate_pdf_paper, generate_pdf_answer_key
        else:
            make_paper, make_key = generate_paper, generate_answer_key
        if job.mode in (MODE_QUESTIONS, MODE_BOTH):
            job.saved.append(make_paper(job.source, job.questions, job.output))
        if job.mode in (MODE_ANSWERS, MODE_BOTH):
            out = job.answer_output or job.output.with_name(
                f"{job.output.stem}_answer_key{job.output.suffix}")
            job.saved.append(make_key(job.source, job.questions, out))

    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(run, job): job for job in jobs}
        for done, future in enumerate(as_completed(futures), 1):
            job = futures[future]
            error = future.exception()
            results.append((job, error))
            if progress_callback:
                progress_callback(done, total, job, error)

    return results
