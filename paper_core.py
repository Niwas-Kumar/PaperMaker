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
from collections import Counter, OrderedDict
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

# Other ways papers number their questions. A document is read in ONE style: the
# one whose numbers form the longest 1, 2, 3 … run ("12." wins a tie, so sub-points
# like "1)" inside a solution never turn a normal paper into this style).
NUMBER_STYLES = (
    ("dot", QUESTION_RE),                                                     # 12.
    ("q", re.compile(r"^\s*(?:Q|Que|Ques|Question)\s*[.\-:]?\s*(\d+)\s*[.):]?")),  # Q.12  Q12.
    ("paren", re.compile(r"^\s*(\d+)\)(?!\d)")),                              # 12)
)


def detect_number_style(texts: Iterable[str]) -> "re.Pattern":
    """The numbering pattern this paper uses for its questions."""
    texts = list(texts)
    best, best_run = QUESTION_RE, -1
    for name, rx in NUMBER_STYLES:
        run = longest = 0
        last = None
        for t in texts:
            m = rx.match(t)
            if not m:
                continue
            n = int(m.group(1))
            run = run + 1 if last is not None and n == last + 1 else 1
            longest = max(longest, run)
            last = n
        if longest > best_run and (name == "dot" or longest >= 5):
            best, best_run = rx, longest
    return best


def _canonical(text: str, style: "re.Pattern") -> str:
    """``Q.12 stem`` / ``12) stem`` -> ``12. stem``: the form the parser reads."""
    if style is QUESTION_RE:
        return text
    m = style.match(text)
    return f"{m.group(1)}. {text[m.end():].lstrip()}" if m else text

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

# A whole answer-key line and nothing else: "46.(3)".
KEY_LINE_RE = re.compile(r"^\s*(\d+)\.\s*\(\s*[^)\s]{1,6}\s*\)\s*$")

# What a missing answer is written as in a generated key: "57.(—)".
MISSING_ANSWER = "(—)"
_KEY_ANSWER_RE = re.compile(r"\(\s*[^)\s]{1,6}\s*\)")

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
    """What the GUI/CLI learn about a source file before building anything.

    *answer_numbers* / *solution_numbers* are the question numbers the file's
    own key and solutions cover. *is_answer_file* is set when the "questions"
    found are really answer-key lines (``46.(3)``): the file is an answer key
    on its own and must be attached to its question paper instead.
    """

    __slots__ = ("subjects", "has_answer_key", "answer_numbers",
                 "solution_numbers", "is_answer_file")

    def __init__(self, subjects, has_answer_key: bool, answer_numbers=(),
                 solution_numbers=(), is_answer_file: bool = False) -> None:
        self.subjects: OrderedDict[str, list[QuestionInfo]] = subjects
        self.has_answer_key = has_answer_key
        self.answer_numbers = frozenset(answer_numbers)
        self.solution_numbers = frozenset(solution_numbers)
        self.is_answer_file = is_answer_file


def parse_paper(path: str | Path) -> ParsedPaper:
    """Parse a .docx into ``ParsedPaper(subjects, has_answer_key, …)``.

    Uses the same segmenter as the builder, so the questions offered in the GUI
    are exactly the blocks that end up in the paper — for both the multi-subject
    (heading-per-subject) layout and single-subject papers with no headings.
    ``has_answer_key`` is true when the source carries an ANSWER KEY part.
    """
    doc = Document(str(path))
    parsed = _segment_questions(doc, _implicit_subject(path))

    subjects: OrderedDict[str, list[QuestionInfo]] = OrderedDict()
    key_lines = 0
    for (subject, number), elements in parsed.segments.items():
        text = ""
        for el in elements:
            if el.tag == qn("w:p"):
                text = Paragraph(el, doc).text.strip()
                break
        key_lines += bool(KEY_LINE_RE.match(text))
        m = next((m for _n, rx in NUMBER_STYLES if (m := rx.match(text))), None)
        raw = text[m.end():].strip() if m else text
        preview = (raw[:77] + "…") if len(raw) > 80 else raw
        subjects.setdefault(subject, []).append(
            QuestionInfo(number=number, subject=subject, preview=preview)
        )
    count = len(parsed.segments)
    is_answer_file = not parsed.answers and count >= 3 and key_lines * 10 >= count * 8
    return ParsedPaper(subjects, bool(parsed.answers), parsed.answers.keys(),
                       parsed.solutions.keys(), is_answer_file)


class ParsedAnswers:
    """What a separate answer-key file covers (see ``parse_answer_file``)."""

    __slots__ = ("answer_numbers", "solution_numbers")

    def __init__(self, answer_numbers, solution_numbers) -> None:
        self.answer_numbers = frozenset(answer_numbers)
        self.solution_numbers = frozenset(solution_numbers)


def parse_answer_file(path: str | Path) -> ParsedAnswers:
    """Read a separate answer-key file: an ANSWER KEY part (``46.(3)`` lines)
    with or without solutions, and no questions.

    Raises ``ValueError`` when no key lines are found, or when the key's
    numbering starts again part-way (e.g. every subject from 1): then answers
    cannot be matched to questions by number.
    """
    doc = Document(str(path))
    parsed = _segment_questions(doc, _implicit_subject(path), answers_only=True)
    if not parsed.answers:
        raise ValueError(
            "No answer key lines (like 46.(3)) were found in this file.")
    bare = sum(1 for els in parsed.solutions.values() if _is_bare_key_block(els, doc))
    if bare >= 5 and bare * 10 >= len(parsed.solutions) * 3:
        raise ValueError(
            "The answer key's numbering starts again part-way (for example each "
            "subject from 1), so answers cannot be matched to the questions.")
    return ParsedAnswers(parsed.answers.keys(), parsed.solutions.keys())


def _is_bare_key_block(elements: list, doc) -> bool:
    """A "solution" that is only a key line: really a piece of the key."""
    content = [el for el in elements if _has_question_content(el)]
    return (len(content) == 1 and content[0].tag == qn("w:p")
            and bool(KEY_LINE_RE.match(Paragraph(content[0], doc).text)))


def parse_document(path: str | Path) -> "OrderedDict[str, list[QuestionInfo]]":
    """Parse a .docx and return ``{subject: [QuestionInfo, …]}``."""
    return parse_paper(path).subjects


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACT & REBUILD ENGINE
# ═══════════════════════════════════════════════════════════════════════════════


class _Parsed:
    """Everything the segmenter harvests from one source document."""

    __slots__ = ("body", "sectPr", "preamble", "heading_templates", "segments",
                 "answer_head", "answers", "solution_head", "solutions",
                 "end_mark", "solution_end", "layout")

    def __init__(self, body, sectPr, preamble, heading_templates, segments,
                 answer_head, answers, solution_head, solutions,
                 end_mark=(), solution_end=(), layout=None) -> None:
        self.body = body
        self.sectPr = sectPr
        self.preamble = preamble
        self.heading_templates = heading_templates
        self.segments = segments
        self.answer_head = answer_head
        self.answers = answers
        self.solution_head = solution_head  # elements between key and solutions
        self.solutions = solutions          # {number: [elements]}
        self.end_mark = list(end_mark)          # closing ornament of the paper
        self.solution_end = list(solution_end)  # … and of the solutions
        self.layout = layout or {}              # {name: _Section} (see _sections_of)


def _carries_section_break(paragraph_el) -> bool:
    pPr = paragraph_el.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:sectPr")) is not None


_REF_TAGS = ("w:headerReference", "w:footerReference")


class _Section:
    """One section of the source: its ``w:sectPr`` and the header/footer
    references in force there (``{(tag, type): rId}``)."""

    __slots__ = ("sectPr", "refs")

    def __init__(self, sectPr, refs) -> None:
        self.sectPr = sectPr
        self.refs = refs


def _sections_of(children: list, body_sectPr, wanted: dict) -> dict:
    """``{name: _Section}`` — the source section each *wanted* element lies in.

    A section's properties are stored at its END (on its last paragraph; on the
    body for the last one), so an element belongs to the first ``w:sectPr`` at
    or after it. The references are resolved the way Word does it: a section
    with no reference of some type keeps the one of the section before.
    """
    names_at: dict[int, list] = {}
    for name, el in wanted.items():
        if el is not None:
            names_at.setdefault(id(el), []).append(name)

    found: dict = {}
    pending: list = []
    refs: dict[tuple[str, str], str] = {}
    for child in children:
        pending += names_at.get(id(child), ())
        if child is body_sectPr:
            sect = child
        elif isinstance(child, CT_P) and _carries_section_break(child):
            sect = child.find(qn("w:pPr")).find(qn("w:sectPr"))
        else:
            continue
        for tag in _REF_TAGS:
            for ref in sect.findall(qn(tag)):
                rid = ref.get(qn("r:id"))
                if rid:
                    refs[(tag, ref.get(qn("w:type")) or "default")] = rid
        for name in pending:
            found[name] = _Section(sect, dict(refs))
        pending = []
    return found


def _has_manual_break(el) -> bool:
    return any(br.get(qn("w:type")) in ("page", "column")
               for br in el.iter(qn("w:br")))


def _has_question_content(el) -> bool:
    """A table, or a paragraph with words, numbers, an equation or a picture —
    as opposed to an empty line, a bare break or a row of ornament symbols."""
    if el.tag == qn("w:tbl"):
        return True
    if el.tag != qn("w:p"):
        return False
    if any(ch.isalnum() for t in el.iter(qn("w:t")) for ch in (t.text or "")):
        return True
    return any(True for _ in el.iter(
        qn("w:drawing"), qn("w:object"), qn("w:pict"), qn("m:oMath")))


def _is_separator(el) -> bool:
    """A plain empty line between two questions (no break of any kind on it)."""
    return (_is_blank_para(el) and not _carries_section_break(el)
            and not _has_manual_break(el))


def _is_end_mark(el) -> bool:
    """A closing ornament line: symbols only, such as ❖❖❖ or * * *."""
    if el.tag != qn("w:p") or _has_question_content(el):
        return False
    if any(True for _ in el.iter(qn("w:sym"))):
        return True
    return bool("".join(t.text or "" for t in el.iter(qn("w:t"))).strip())


def _tidy_blocks(blocks: "OrderedDict") -> list:
    """Leave every block with the paper's usual blank line(s) after it.

    In the source, what follows a question's last line is not always its own:
    the blank line is deleted where a question ends a column, and the last
    question of a subject or of the paper is followed by spare blank lines,
    section breaks, the page break before the answer key or the closing
    ornament. Moved to another place those glue two questions together, open a
    hole or break the page in the middle of the paper. So each block ends with
    as many blank lines as most blocks of this paper have — its own, or a copy
    of a neighbour's — and nothing else.

    Returns the end mark: the ornament line(s) that follow the last block, with
    the blank lines leading to them, for the builder to close the new paper with.
    """
    cuts = {}
    for key, els in blocks.items():
        cut = len(els)
        while cut and not _has_question_content(els[cut - 1]):
            cut -= 1
        if cut:
            cuts[key] = cut
    if not cuts:
        return []

    seps = {key: [el for el in blocks[key][cut:] if _is_separator(el)]
            for key, cut in cuts.items()}
    usual, votes = Counter(len(s) for s in seps.values()).most_common(1)[0]
    settled = votes * 10 >= len(cuts) * 7  # a real habit of this paper
    plain = [el for s in seps.values() for el in s[:1]]
    spare = next((el for el in plain  # one with nothing but its formatting
                  if all(c.tag == qn("w:pPr") for c in el)), plain[0] if plain else None)
    last_key = next(reversed(blocks))

    end_mark: list = []
    for key, cut in cuts.items():
        els = blocks[key]
        tail = els[cut:]
        paras = [el for el in tail if el.tag == qn("w:p")]
        if not settled and all(_is_separator(el) for el in paras):
            continue  # no clear habit and nothing foreign: leave it as typed
        keep = seps[key][:usual]
        if settled and spare is not None:
            for _ in range(usual - len(keep)):
                blank = copy.deepcopy(spare)
                for name in [a for a in blank.attrib if a.endswith(("}paraId", "}textId"))]:
                    del blank.attrib[name]  # ids must stay unique in the document
                keep.append(blank)
        marks = [i for i, el in enumerate(tail) if _is_end_mark(el)]
        if key == last_key and marks:
            end_mark = [el for el in tail[:marks[-1] + 1]
                        if (_is_separator(el) or _is_end_mark(el))
                        and not any(el is k for k in keep)]
        els[cut:] = [el for el in tail if el.tag != qn("w:p")] + keep
    return end_mark


def _strip_manual_breaks(el) -> None:
    """Drop page/column breaks typed into a block: they fitted the old page
    and would force a new column or page in the middle of the new one."""
    for br in list(el.iter(qn("w:br"))):
        if br.get(qn("w:type")) in ("page", "column"):
            br.getparent().remove(br)
    for node in list(el.iter(qn("w:pageBreakBefore"))):
        node.getparent().remove(node)


def _segment_questions(
    doc: DocumentObject,
    implicit_subject: str = "Questions",
    answers_only: bool = False,
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

    *answers_only* reads a separate answer-key file, which has no questions:
    everything before the marker (or before the first key line, when there is
    no marker) is the preamble. Such a key may cover several subjects, each in
    its own section, so a section break does not end it — only the numbering
    starting again (the solutions) does; what lies between the key and the
    solutions becomes the solutions' head, as in a combined paper. Subject
    banners inside the solutions are left out of the blocks.
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
    style = detect_number_style(
        Paragraph(c, doc).text for c in children if isinstance(c, CT_P))

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
    gap: list | None = None  # answers_only: what follows a section break in the key
    # "questions" → "answer_head" → "answers" → "solutions" → "done";
    # an answer-key file starts at "lead" instead of "questions"
    phase = "lead" if answers_only else "questions"

    for child in children:
        if child is sectPr or phase == "done":
            continue
        is_para = isinstance(child, CT_P)
        text = _canonical(Paragraph(child, doc).text.strip(), style) if is_para else ""

        if phase == "lead":
            if is_para and ANSWER_MARKER_RE.match(text):
                answer_head.append(child)
                phase = "answer_head"
                continue
            if not (is_para and SOLUTION_START_RE.match(text)):
                preamble.append(child)
                continue
            phase = "answers"  # a key with no heading of its own — fall through

        if phase == "answer_head":
            if is_para and QUESTION_RE.match(text):
                phase = "answers"  # first answer line — fall through
            else:
                answer_head.append(child)
                continue

        if phase == "answers" and answers_only:
            m = QUESTION_RE.match(text) if is_para else None
            if m and last_answer is not None and int(m.group(1)) <= last_answer:
                phase = "solutions"  # numbering restarted: this IS a solution
                solution_head.extend(gap or ())
            elif gap is not None and not m:
                gap.append(child)  # between the key's parts, or before solutions
                continue
            else:
                if m:
                    answers.setdefault(int(m.group(1)), child)
                    last_answer = int(m.group(1))
                    gap = None
                if is_para and _carries_section_break(child):
                    gap = []
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
            if answers_only and is_para and heading_of.get(id(child)):
                continue  # a subject banner between two subjects' solutions
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

    # Where each part sat in the source (read before any block is touched).
    wanted: dict = {"body": sectPr}
    for (subject, _number), elements in segments.items():
        if elements:
            wanted.setdefault(("q", subject), elements[0])
    for subject, heading in heading_templates.items():
        wanted[("h", subject)] = heading
    if answers:
        wanted["key"] = next(iter(answers.values()))
    wanted["solutions"] = next((els[0] for els in solutions.values() if els), None)
    layout = _sections_of(children, sectPr, wanted)

    end_mark = _tidy_blocks(segments)
    solution_end = _tidy_blocks(solutions)
    return _Parsed(body, sectPr, preamble, heading_templates, segments,
                   answer_head, answers, solution_head, solutions,
                   end_mark, solution_end, layout)


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
    """Set the column count of *sectPr* (``num`` = 1 for a full-width section).

    The gap between the columns stays the source's own; 709 is only the
    default for a section that never had one.
    """
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sectPr.append(cols)
    if cols.get(qn("w:num"), "1") != str(max(1, num)):
        for col in list(cols):  # per-column widths of the old count
            cols.remove(col)
        cols.attrib.pop(qn("w:equalWidth"), None)
    cols.set(qn("w:num"), str(max(1, num)))
    if cols.get(qn("w:space")) is None:
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


def _fresh_sectpr(base_sectPr, keep_title_pg: bool = False) -> object:
    """Deep-copy *base_sectPr* (page size/margins/columns) stripped of layout.

    Header/footer refs, the break type, a page-number restart and ``titlePg``
    are removed so callers can set exactly the layout they need; the page
    geometry is preserved.
    """
    sect = copy.deepcopy(base_sectPr)
    drop = ["w:headerReference", "w:footerReference", "w:type", "w:pgNumType"]
    if not keep_title_pg:
        drop.append("w:titlePg")
    for tag in drop:
        for el in sect.findall(qn(tag)):
            sect.remove(el)
    return sect


def _rebuilt_sectpr(
    section: "_Section | None",
    body_sectPr,
    *,
    cols: int | None,
    footer_refs: list[tuple[str, str]],
    own_refs: bool,
) -> object:
    """Section properties for a rebuilt run of content (a continuous section).

    Page size, margins and column gap are those of the source section the
    content sat in (*section*; the body's when unknown), so lines wrap where
    they did. *cols* forces a column count, ``None`` keeps the source's. With
    *own_refs* the section shows the headers and footers it had in the source
    (papers with a running header); otherwise only the page-number footer is
    attached, like the content pages of a paper with a cover.
    """
    base = section.sectPr if section is not None else body_sectPr
    sect = _fresh_sectpr(base, keep_title_pg=own_refs and section is not None)
    if cols is not None:
        _set_cols(sect, cols)
    _set_type(sect, "continuous")
    if own_refs and section is not None:
        for (tag, type_val), rid in section.refs.items():
            _ensure_ref(sect, tag, type_val, rid)
    else:
        _attach_footer_refs(sect, footer_refs)
    return sect


def _section_break_para(sect) -> object:
    """Build an empty paragraph that ends a section with properties *sect*."""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    p.append(pPr)
    pPr.append(sect)
    return p


def _heading_para(template, subject: str, sect) -> object:
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
            other = None if m else next(
                (o for _n, rx in NUMBER_STYLES[1:] if (o := rx.match(joined))), None)
            if other and (i + 1 == len(texts) or other.end() < len(joined)
                          or not joined[-1].isdigit()):
                # "Q.12" / "12)": only the digits change, the paper's style stays
                end = other.end()
                head = (joined[:other.start(1)] + str(new_number)
                        + joined[other.end(1):end])
                pos = 0
                for k, part in enumerate(texts[:i + 1]):
                    old = part.text or ""
                    part.text = (head if k == 0 else "") + old[max(0, end - pos):]
                    pos += len(old)
                return

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


def _switch_on(paragraph_el, tag: str) -> None:
    """Turn on ``w:keepNext`` / ``w:keepLines`` in schema order (after pStyle)."""
    pPr = paragraph_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        paragraph_el.insert(0, pPr)
    el = pPr.find(qn(tag))
    if el is not None:
        el.attrib.pop(qn("w:val"), None)
        return
    el = OxmlElement(tag)
    before = ("w:pStyle", "w:keepNext") if tag == "w:keepLines" else ("w:pStyle",)
    anchor = next((a for a in (pPr.find(qn(t)) for t in reversed(before))
                   if a is not None), None)
    if anchor is not None:
        anchor.addnext(el)
    else:
        pPr.insert(0, el)


def _is_blank_para(el) -> bool:
    if el.tag != qn("w:p"):
        return False
    if "".join(t.text or "" for t in el.iter(qn("w:t"))).strip():
        return False
    return not any(True for _ in el.iter(
        qn("w:drawing"), qn("w:object"), qn("w:pict"), qn("m:oMath"), qn("w:sym")))


def keep_block_together(elements: list) -> None:
    """Keep one question — stem, figures, tables and options — in one column.

    Every paragraph keeps its lines together and stays with the next one, and
    table rows may not split, so Word moves the whole question to the next
    column or page instead of cutting it. The blank line after the question
    travels with it (so it never opens the next column); only that last line
    is free, and the next question can start in a new column. A question
    taller than a whole column still has to break.
    """
    block = [el for el in elements if el.tag in (qn("w:p"), qn("w:tbl"))]
    for i, el in enumerate(block):
        last = i == len(block) - 1
        if el.tag == qn("w:p"):
            _switch_on(el, "w:keepLines")
            if not last:
                _switch_on(el, "w:keepNext")
            continue
        rows = el.findall(qn("w:tr"))
        for r, tr in enumerate(rows):
            trPr = tr.find(qn("w:trPr"))
            if trPr is None:
                trPr = OxmlElement("w:trPr")
                at = next((k for k, c in enumerate(tr)
                           if c.tag != qn("w:tblPrEx")), len(tr))
                tr.insert(at, trPr)
            if trPr.find(qn("w:cantSplit")) is None:
                cant_split = OxmlElement("w:cantSplit")
                later = next((trPr.find(qn(t)) for t in (
                    "w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc",
                    "w:hidden", "w:ins", "w:del", "w:trPrChange")
                    if trPr.find(qn(t)) is not None), None)
                if later is not None:
                    later.addprevious(cant_split)
                else:
                    trPr.append(cant_split)
            if last and r == len(rows) - 1:
                continue
            for p in tr.iter(qn("w:p")):
                _switch_on(p, "w:keepNext")


def _join_names(names: list[str]) -> str:
    """Join subject names as ``A``, ``A and B`` or ``A, B and C``."""
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    return f"{', '.join(names[:-1])} and {names[-1]}"


BIOLOGY_PARTS = ("botany", "zoology")


def _subject_selected(name: str, selected_lower: set[str]) -> bool:
    """Whether a cover line about subject *name* still applies to this paper.

    A paper whose banner reads BIOLOGY lists Botany and Zoology on its cover
    (and the other way round), so those stand or fall together.
    """
    if name in selected_lower:
        return True
    if name == "biology":
        return any(part in selected_lower for part in BIOLOGY_PARTS)
    return name in BIOLOGY_PARTS and "biology" in selected_lower


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
        if row_subj is not None and not _subject_selected(row_subj, selected_lower):
            tbl.remove(tr)


# Cover-page figures that follow the selection; group 1 is the part rewritten.
_COUNT_RE = re.compile(r"(\d+)(?=\s+multiple[\s-]*choice\s+questions?)", re.I)
_MARKS_RE = re.compile(r"max(?:imum|\.)?\s+marks?\s*(?:is|are|[:=–-])?\s*(\d+)", re.I)
_SUBJECTS_RE = re.compile(r"\bfrom\s+([^.]*?)\s*\.\s*All questions", re.I)


def _run_texts(paragraph_el) -> list:
    return paragraph_el.findall(qn("w:r") + "/" + qn("w:t"))


def _replace_span(paragraph_el, start: int, end: int, new_text: str) -> None:
    """Replace characters ``[start, end)`` of the paragraph's text.

    The span may run over several runs (a bold figure, a bracket typed later);
    the new text takes the formatting of the run the span starts in.
    """
    pos, written = 0, False
    for t in _run_texts(paragraph_el):
        old = t.text or ""
        a, b = pos, pos + len(old)
        pos = b
        if b <= start or a >= end:
            continue
        text = old[:max(0, start - a)]
        if not written:
            text += new_text
            written = True
        if end < b:
            text += old[end - a:]
        t.text = text
        if text != text.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _customize_preamble(
    preamble: list,
    subjects_order: list[str],
    total: int,
    all_subjects: Iterable[str] = (),
    source_total: int = 180,
) -> None:
    """Rewrite the cover page's instruction counts and syllabus for this paper.

    The question total, the maximum mark (``4 × total``) and the subject list are
    all made to reflect the questions the teacher actually selected. The subject
    list is left exactly as printed when every subject of the source is there.
    """
    marks = total * 4
    old_marks = str(source_total * 4)
    phrase = _join_names([s for s in subjects_order])
    selected_lower = {s.lower() for s in subjects_order}
    subjects_changed = selected_lower != {s.lower() for s in all_subjects}

    for el in preamble:
        if el.tag == qn("w:tbl"):
            _trim_syllabus(el, selected_lower)
            continue
        if el.tag != qn("w:p"):
            continue
        rewrites = [(_COUNT_RE, str(total)), (_MARKS_RE, str(marks))]
        if phrase and subjects_changed:
            rewrites.append((_SUBJECTS_RE, phrase))
        for pattern, value in rewrites:
            # the sentence may be cut into runs anywhere: search the whole line
            m = pattern.search("".join(t.text or "" for t in _run_texts(el)))
            if not m:
                continue
            if pattern is _SUBJECTS_RE and not any(
                    kw.lower() in m.group(1).lower() for kw in SUBJECT_KEYWORDS):
                continue
            _replace_span(el, m.start(1), m.end(1), value)
        if "mark" in "".join(t.text or "" for t in _run_texts(el)).lower():
            for t in _run_texts(el):
                if t.text and t.text.strip() == old_marks:  # the figure on its own
                    t.text = t.text.replace(old_marks, str(marks))


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
    questions ``1, 2, 3, …`` in the teacher's chosen order (a subject's
    questions stay together under its one banner). Each subject banner sits in
    its own single-column section and each question block in a section with the
    page margins and columns its questions had in the source, so lines wrap
    where they did; the paper closes with the source's own end mark.
    """
    doc = Document(str(source))
    _, footer_refs = _discover_layout_refs(doc)
    parsed = _segment_questions(doc, _implicit_subject(source))
    body, sectPr, preamble = parsed.body, parsed.sectPr, parsed.preamble
    heading_templates, segments = parsed.heading_templates, parsed.segments
    # Papers without subject headings (single-subject tests) get no cover
    # rewrite and no heading bars: their own first page is mirrored as-is.
    has_headings = bool(heading_templates)
    source_total = len(segments)

    picked = [(subj, els) for subj, els in pick(segments) if els]
    picked = group_by_subject(picked, subject=lambda pair: pair[0])

    subjects_order: list[str] = []
    for subj, _els in picked:
        if subj not in subjects_order:
            subjects_order.append(subj)
    total = len(picked)

    def section(kind: str, subj: str, cols: int | None):
        """Section properties for *subj*'s banner ("h") or questions ("q")."""
        known = parsed.layout.get((kind, subj))
        if known is None:
            known, cols = parsed.layout.get("body"), cols or 2
        return _rebuilt_sectpr(known, sectPr, cols=cols, footer_refs=footer_refs,
                               own_refs=not has_headings)

    for child in list(body.iterchildren()):
        if child is not sectPr:
            body.remove(child)

    # Cover page (first single-column section with its own footer / title page),
    # reused verbatim so the title box, logo and spacing match the source exactly.
    if preamble:
        if has_headings:
            _customize_preamble(preamble, subjects_order, total,
                                heading_templates, source_total)
        for el in preamble:
            sectPr.addprevious(el)

    prev_subject: str | None = None
    number = 0
    for subj, element_list in picked:
        if subj != prev_subject and has_headings:
            if prev_subject is not None:
                # Close the previous subject's question section.
                sectPr.addprevious(_section_break_para(
                    section("q", prev_subject, None)))
            sectPr.addprevious(_heading_para(
                heading_templates.get(subj), subj, section("h", subj, 1)))
        prev_subject = subj

        number += 1
        _renumber_segment(element_list, number)
        kept = []
        for el in element_list:
            if el.tag == qn("w:p") and _strip_section_break(el):
                continue
            _strip_manual_breaks(el)
            sectPr.addprevious(el)
            kept.append(el)
        # A question and its options never split across columns or pages.
        keep_block_together(kept)

    if picked:
        end_mark = [el for el in parsed.end_mark if not _strip_section_break(el)]
        for el in end_mark:
            sectPr.addprevious(el)
        if end_mark:
            # the ornament never stands alone on top of a new column or page
            keep_block_together(kept + end_mark)
        _drop_trailing_blanks(sectPr)
        # The body's own (terminal) section closes the last subject's block.
        # No break is put after the last question: one would make Word level
        # the two columns of the last page, and a viewer that levels without
        # honouring "keep with next" would cut a question doing it.
        body.replace(sectPr, section("q", prev_subject, None))
    _drop_unused_media(doc)
    return _save_with_fallback(doc, output)


def _drop_trailing_blanks(body_sectPr) -> None:
    """Remove the empty lines at the very end of the rebuilt body.

    Nothing follows them, and when the last block ends exactly at the foot of a
    page they would be alone on one more, otherwise empty, page.
    """
    last = body_sectPr.getprevious()
    while last is not None and _is_separator(last):
        last.getparent().remove(last)
        last = body_sectPr.getprevious()


def group_by_subject(items: Iterable, subject=lambda q: q.subject) -> list:
    """*items* with each subject's entries brought together.

    Subjects come in the order they first appear and the order inside a subject
    is kept, so a paper never shows the same subject banner twice.
    """
    groups: "OrderedDict[str, list]" = OrderedDict()
    for item in items:
        groups.setdefault(subject(item), []).append(item)
    return [item for group in groups.values() for item in group]


_UNUSED_OK = ("/image", "/oleObject", "/package")


def _drop_unused_media(doc: DocumentObject) -> None:
    """Leave out pictures and equation objects of questions that were not kept.

    The source's relationships still list every embedded file, so without this
    a five-question paper would carry (and Word would load and virus-scan) the
    media of all 180. Only image / embedded-object relationships that nothing
    in the rebuilt body mentions are dropped.
    """
    rels = doc.part.rels
    used = {value for el in doc.element.iter()
            for value in el.attrib.values() if value in rels}
    for rid, rel in list(rels.items()):
        if rid in used or rel.is_external:
            continue
        if rel.reltype.endswith(_UNUSED_OK):
            del rels[rid]
            rels.related_parts.pop(rid, None)


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


def _copy_without_ids(el):
    """A deep copy of paragraph *el* that can sit next to the original."""
    p = copy.deepcopy(el)
    for name in [a for a in p.attrib if a.endswith(("}paraId", "}textId"))]:
        del p.attrib[name]  # ids must stay unique in the document
    return p


def _key_line(answer_el, number: int, missing: bool = False):
    """A copy of key line *answer_el* (``46.(3)``) renumbered to *number*;
    with *missing* its answer reads ``(—)`` — what the key shows for a
    question the source has no answer for."""
    p = _copy_without_ids(answer_el)
    _strip_section_break(p)
    _renumber_segment([p], number)
    if missing:
        texts = p.findall(qn("w:r") + "/" + qn("w:t"))
        hit = next((t for t in texts if _KEY_ANSWER_RE.search(t.text or "")), None)
        if hit is not None:  # keep the line's own tab / spacing
            hit.text = _KEY_ANSWER_RE.sub(MISSING_ANSWER, hit.text, count=1)
        else:  # the answer is split over runs: one plain text instead
            for i, t in enumerate(texts):
                t.text = f"{number}.{MISSING_ANSWER}" if i == 0 else ""
    return p


def generate_answer_key(
    source_path: str | Path,
    selected_questions: list[QuestionInfo],
    output_path: str | Path,
    *,
    answer_file: bool = False,
) -> Path:
    """Build a separate answer-key paper for *selected_questions*.

    The source's own ANSWER KEY part supplies the answer lines; the output keeps
    the source's first page and the key's own heading/section layout, and the
    answers are renumbered to match the question paper built from the same
    selection (``1, 2, 3, …`` in the chosen order). A question the key has no
    line for gets ``N.(—)``; with solutions, one with no solution gets just its
    key line. With *answer_file* the source is a separate answer-key file
    (see ``parse_answer_file``) and *selected_questions* come from its question
    paper. Raises ``ValueError`` when the source has no answer key. Returns the
    path actually written.
    """
    if not selected_questions:
        raise ValueError("No questions selected.")
    source = Path(source_path)
    output = Path(output_path)

    doc = Document(str(source))
    _, footer_refs = _discover_layout_refs(doc)
    parsed = _segment_questions(doc, _implicit_subject(source), answers_only=answer_file)
    if not parsed.answers:
        raise ValueError("No answer key was detected in this paper.")
    template = next(iter(parsed.answers.values()))

    body, sectPr = parsed.body, parsed.sectPr
    own_refs = not parsed.heading_templates

    def section(name: str):
        """Section properties of the key ("key") or the solutions part."""
        known, cols = parsed.layout.get(name), None
        if known is None:
            known, cols = parsed.layout.get("body"), 2
        return _rebuilt_sectpr(known, sectPr, cols=cols, footer_refs=footer_refs,
                               own_refs=own_refs)

    for child in list(body.iterchildren()):
        if child is not sectPr:
            body.remove(child)
    for el in parsed.preamble:
        sectPr.addprevious(el)
    for el in parsed.answer_head:
        sectPr.addprevious(el)

    # The paper's numbering: a selection missing from the paper consumes none.
    # (An answer file has no questions; the selection comes from its paper.)
    numbered: list[tuple[int, QuestionInfo]] = []
    for qi in group_by_subject(selected_questions):
        if answer_file or parsed.segments.get((qi.subject, qi.number)):
            numbered.append((len(numbered) + 1, qi))

    for number, qi in numbered:
        answer = parsed.answers.get(qi.number)
        if answer is None:  # question exists but the key has no line for it
            sectPr.addprevious(_key_line(template, number, missing=True))
            continue
        answer = copy.deepcopy(answer)
        _strip_section_break(answer)
        _renumber_segment([answer], number)
        sectPr.addprevious(answer)

    # Solutions, when the source has them: close the key's columns the way the
    # source does, keep the source's own gap, then the selected blocks under
    # the same new numbers.
    last_part = "key"
    if parsed.solutions:
        sectPr.addprevious(_section_break_para(section("key")))
        for el in parsed.solution_head:
            sectPr.addprevious(el)
        blank = next((els[-1] for els in parsed.solutions.values()
                      if els and _is_separator(els[-1])), None)
        for number, qi in numbered:
            block = parsed.solutions.get(qi.number)
            if not block:  # no solution: its key line alone, then the usual gap
                answer = parsed.answers.get(qi.number)
                sectPr.addprevious(_key_line(template, number, missing=True)
                                   if answer is None else _key_line(answer, number))
                if blank is not None:
                    sectPr.addprevious(_copy_without_ids(blank))
                continue
            _renumber_segment(block, number)
            for el in block:
                if el.tag == qn("w:p") and _strip_section_break(el):
                    continue
                _strip_manual_breaks(el)
                sectPr.addprevious(el)
        for el in parsed.solution_end:
            if not _strip_section_break(el):
                sectPr.addprevious(el)
        last_part = "solutions"
    _drop_trailing_blanks(sectPr)
    # The body's own (terminal) section closes the last part's columns.
    body.replace(sectPr, section(last_part))
    _drop_unused_media(doc)
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

    __slots__ = ("source", "questions", "output", "mode", "answer_output", "saved",
                 "answer_source")

    def __init__(
        self,
        source: str | Path,
        questions: list[QuestionInfo],
        output: str | Path,
        mode: str = MODE_QUESTIONS,
        answer_output: str | Path | None = None,
        answer_source: str | Path | None = None,
    ) -> None:
        self.source = Path(source)
        self.questions = questions
        self.output = Path(output)
        self.mode = mode
        self.answer_output = Path(answer_output) if answer_output else None
        # a separate answer-key .docx attached to the paper (else: its own key)
        self.answer_source = Path(answer_source) if answer_source else None
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
            if job.answer_source is not None:
                job.saved.append(generate_answer_key(
                    job.answer_source, job.questions, out, answer_file=True))
            else:
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
