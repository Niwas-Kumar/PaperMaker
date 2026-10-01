#!/usr/bin/env python3
"""PDF source support for Paper Maker.

A PDF has no editable structure, so questions are located from its text layer
(PyMuPDF) and rebuilt into a Word file:

* a question that is plain text becomes real, editable Word text (lines are
  re-clustered by baseline, options kept one per line or side by side, super-
  and subscripts restored);
* a question containing a photo, a drawing (graph) or stacked maths (fractions)
  is placed as a 300-dpi picture cut from the original, so it looks identical.

The answer key (``ANSWER KEY`` + ``N.(k)`` lines) and the solutions that follow
it are harvested the same way for the companion answer-key file.
"""

from __future__ import annotations

import io
import re
from collections import OrderedDict
from pathlib import Path

import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Mm, Pt

from paper_core import (
    ANSWER_MARKER_RE,
    ParsedPaper,
    QuestionInfo,
    _implicit_subject,
    _save_with_fallback,
)

# Question number at the start of a line ("136." / "*9."); the lookahead keeps
# "2.5 m/s" from being read as question 2.
QRE = re.compile(r"^\s*\*?\s*(\d{1,3})\.(?!\d)")
# Answer-key / solution line: number AND answer, "136. (3)".
KEY_RE = re.compile(r"^\s*\*?\s*(\d{1,3})\.\s*\(\s*([^)\s]{1,6})\s*\)")
OPT_RE = re.compile(r"\(\d\)")
LABEL_RE = re.compile(
    r"^(Assertion|Reason|Statement|List|Column|Match|Choose|Which|[A-E]\.)\b")

BODY_PT = 10          # body font size of the generated Word text
PICTURE_DPI = 300


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════


class _Line:
    __slots__ = ("col", "y0", "y1", "x0", "x1", "spans", "text")

    def __init__(self, col, y0, y1, x0, x1, spans):
        self.col, self.y0, self.y1, self.x0, self.x1 = col, y0, y1, x0, x1
        self.spans = spans
        self.text = _join_spans(spans)


# Symbol-font glyphs that PyMuPDF reports as private-use code points.
_SYMBOL_MAP = str.maketrans({
    0xF0B4: "×", 0xF02D: "−", 0xF0B0: "°", 0xF0B1: "±", 0xF0A3: "≤", 0xF0B3: "≥",
    0xF0B9: "≠", 0xF0BB: "≈", 0xF0AC: "←", 0xF0AE: "→", 0xF0A5: "∞", 0xF0D6: "√",
    0xF0E5: "∑", 0xF0F2: "∫", 0xF0B7: "·", 0xF0A8: "÷", 0xF0B6: "∂", 0xF0CE: "∈",
    0xF0B5: "∝", 0xF0D1: "∇", 0xF0A2: "′", 0xF0D7: "⋅", 0xF0BC: "…", 0xF0DE: "⇒",
    0xF0DB: "⇔", 0xF0A0: "€", 0xF0AB: "↔", 0xF0AD: "↑", 0xF0AF: "↓", 0xF0A6: "ƒ",
    0xF0E6: "(", 0xF0F6: ")", 0xF0E7: "(", 0xF0F7: ")", 0xF0E8: "(", 0xF0F8: ")",
    0xF0E9: "[", 0xF0F9: "]", 0xF0EA: "[", 0xF0FA: "]", 0xF0EB: "[", 0xF0FB: "]",
})


def _span_text(sp) -> str:
    return sp["text"].translate(_SYMBOL_MAP)


def _join_spans(spans) -> str:
    out = ""
    last_x1 = None
    for sp in spans:
        t = _span_text(sp)
        if last_x1 is not None and sp["bbox"][0] - last_x1 > 1 and out and \
                not out.endswith(" ") and not t.startswith(" "):
            out += " "
        out += t
        last_x1 = sp["bbox"][2]
    return re.sub(r"[ \t]{2,}", " ", out).strip()


def _calibrate(doc) -> tuple[float, float]:
    """Header/footer band heights from text that repeats at the same place on
    consecutive pages (a running header/footer does; body text does not)."""
    H = doc[0].rect.height

    def lines(page):
        return [(round(l["bbox"][1]), l["bbox"][3],
                 "".join(s["text"] for s in l["spans"]).strip())
                for b in page.get_text("dict")["blocks"] if b["type"] == 0
                for l in b["lines"]]

    if doc.page_count >= 3:
        a, b = lines(doc[1]), lines(doc[2])
        keyb = {(y0, t) for y0, _y1, t in b if t}
        rep = [(y0, y1) for y0, y1, t in a if t and (y0, t) in keyb]
        head = [y1 for y0, y1 in rep if y0 < H * 0.15]
        header = (max(head) + 4) if head else 60
        foot_ys = {y0 for y0, _y1 in rep if y0 > H * 0.88}
        foot = [y0 for y0, _y1, t in a
                if t and any(abs(y0 - fy) <= 1 for fy in foot_ys)]
        footer = (H - min(foot) + 3) if foot else 18
    else:
        ys = lines(doc[0])
        head = [y1 for y0, y1, t in ys if t and y0 < 60]
        header = (max(head) + 4) if head else 60
        footer = 18
    return header, footer


def _column_split(W: float, raw) -> float:
    """x of the gutter: centre of the widest text-free band near mid-page."""
    xs = sorted((r[0], r[2]) for r in raw if r[2] - r[0] < W * 0.6)
    best, bx, reach = 0.0, W / 2, 0.0
    for x0, x1 in xs:
        if x0 - reach > best and W * 0.3 < (x0 + reach) / 2 < W * 0.7:
            best, bx = x0 - reach, (x0 + reach) / 2
        reach = max(reach, x1)
    return bx


class _Page:
    """Body lines of one page in column-major reading order, plus geometry."""

    def __init__(self, doc, pno: int, header_h: float, footer_h: float):
        page = doc[pno]
        self.pno = pno
        self.W, self.H = page.rect.width, page.rect.height
        d = page.get_text("dict")
        raw = []
        for b in d["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                x0, y0, x1, y1 = l["bbox"]
                if y0 < header_h or y1 > self.H - footer_h:
                    continue
                spans = [s for s in l["spans"] if s["text"].strip()]
                if spans:
                    raw.append([x0, y0, x1, y1, spans])
        self.split = _column_split(self.W, raw)
        col_of = lambda r: 0 if (r[0] + r[2]) / 2 < self.split else 1
        raw.sort(key=lambda r: (col_of(r), r[1], r[0]))
        merged: list[_Line] = []
        for r in raw:
            col = col_of(r)
            # fragments of one printed line (justified text, "1." + "(3)") share a top
            if merged and merged[-1].col == col and abs(merged[-1].y0 - r[1]) < 1.5:
                m = merged[-1]
                m.x0, m.x1, m.y1 = min(m.x0, r[0]), max(m.x1, r[2]), max(m.y1, r[3])
                m.spans = sorted(m.spans + r[4], key=lambda s: s["bbox"][0])
                m.text = _join_spans(m.spans)
            else:
                merged.append(_Line(col, r[1], r[3], r[0], r[2], r[4]))
        self.lines = merged
        self.images = [(0 if (b["bbox"][0] + b["bbox"][2]) / 2 < self.split else 1,
                        fitz.Rect(b["bbox"]))
                       for b in d["blocks"] if b["type"] == 1 and b["bbox"][1] >= header_h]
        self.extents: dict[int, list[float]] = {}
        for m in merged:
            e = self.extents.setdefault(m.col, [m.y0, m.y1, m.x0, m.x1])
            e[0], e[1] = min(e[0], m.y0), max(e[1], m.y1)
            e[2], e[3] = min(e[2], m.x0), max(e[3], m.x1)
        for col, ib in self.images:
            e = self.extents.setdefault(col, [ib.y0, ib.y1, ib.x0, ib.x1])
            e[0], e[1] = min(e[0], ib.y0), max(e[1], ib.y1)


class _Anchor:
    """Where a question / solution block starts."""

    __slots__ = ("number", "pno", "col", "y", "nbox", "size", "x_stem")

    def __init__(self, number, pno, col, y, nbox, size, x_stem):
        self.number, self.pno, self.col, self.y = number, pno, col, y
        self.nbox, self.size, self.x_stem = nbox, size, x_stem


def _make_anchor(pg: _Page, ln: _Line, number: int, prev_bottom: float) -> _Anchor:
    num_span = ln.spans[0]
    size = num_span["size"]
    words = [sp for sp in ln.spans[1:] if re.search(r"[A-Za-z0-9(]", sp["text"])]
    x_stem = words[0]["bbox"][0] if words else num_span["bbox"][2] + 2
    start_y = ln.y0
    # a picture between the previous text line and this number is this block's
    for col, ib in pg.images:
        if col == ln.col and ib.y0 >= prev_bottom - 2 and ib.y0 < ln.y1:
            start_y = min(start_y, ib.y0)
    # so is an equation centred on the number line (stacked rows just above it)
    rows = sorted({round(sp["origin"][1], 1) for l2 in pg.lines
                   if l2.col == ln.col and ln.y0 - 3 * size < l2.y0 < ln.y0
                   and l2.x0 > ln.spans[0]["bbox"][2] - 2
                   for sp in l2.spans if sp["size"] >= size * 0.85})
    if any(size * 0.3 < b - a < size * 0.85 for a, b in zip(rows, rows[1:])):
        start_y = min(start_y, rows[0] - size)
    return _Anchor(number, pg.pno, ln.col, start_y, fitz.Rect(num_span["bbox"]),
                   size, x_stem)


class PdfPaper:
    """Everything located in one PDF."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.doc = fitz.open(str(self.path))
        self.header_h, self.footer_h = _calibrate(self.doc)
        self.pages = [_Page(self.doc, i, self.header_h, self.footer_h)
                      for i in range(self.doc.page_count)]
        self.questions: list[_Anchor] = []
        self.answers: "OrderedDict[int, str]" = OrderedDict()
        self.solutions: list[_Anchor] = []
        self.marker: tuple[int, int, float] | None = None
        self._walk()
        if self.questions:
            first = self.questions[0]
            for col, e in self.pages[first.pno].extents.items():
                e[0] = max(e[0], first.y)  # cover lines above Q1 are not body
        self.q_regions = self._regions(self.questions, self.marker)
        self.s_regions = self._regions(self.solutions, None)

    # ── locating ──────────────────────────────────────────────────────────
    def _walk(self) -> None:
        self._phase = "questions"
        self._last = {"q": None, "a": None, "s": None}
        for pg in self.pages:
            prev_bottom = {0: 0.0, 1: 0.0}
            # The "ANSWER KEY" title is a centred, full-width line: everything on
            # its page above it is still questions, everything below it is the
            # key — regardless of which column the title's x falls into.
            marker_ln = None
            if self._phase == "questions":
                marker_ln = next((ln for ln in pg.lines
                                  if ANSWER_MARKER_RE.match(ln.text)), None)
            if marker_ln is None:
                self._walk_lines(pg, pg.lines, prev_bottom)
                continue
            self._walk_lines(pg, [ln for ln in pg.lines if ln.y0 < marker_ln.y0 - 1],
                             prev_bottom)
            self.marker = (pg.pno, 0, marker_ln.y0)
            self._phase = "answers"
            self._walk_lines(pg, [ln for ln in pg.lines
                                  if ln.y0 >= marker_ln.y0 - 1 and ln is not marker_ln],
                             prev_bottom)

    def _walk_lines(self, pg: _Page, lines: list[_Line], prev_bottom: dict) -> None:
        last = self._last
        for ln in lines:
            t = ln.text
            if self._phase == "questions":
                m = QRE.match(t)
                if m and (last["q"] is None or int(m.group(1)) > last["q"]):
                    last["q"] = int(m.group(1))
                    self.questions.append(
                        _make_anchor(pg, ln, last["q"], prev_bottom[ln.col]))
            elif self._phase == "answers":
                m = KEY_RE.match(t)
                if m:
                    n = int(m.group(1))
                    if last["a"] is not None and n <= last["a"]:
                        self._phase = "solutions"  # numbering restarted
                    else:
                        self.answers[n] = f"({m.group(2)})"
                        last["a"] = n
            if self._phase == "solutions":
                m = KEY_RE.match(t)
                if m and (last["s"] is None or int(m.group(1)) > last["s"]):
                    last["s"] = int(m.group(1))
                    self.solutions.append(
                        _make_anchor(pg, ln, last["s"], prev_bottom[ln.col]))
            prev_bottom[ln.col] = ln.y1

    def _regions(self, anchors: list[_Anchor], end) -> dict[int, list]:
        """number -> [(pno, rect), …] covering the block, column by column."""
        keys = sorted((pg.pno, col) for pg in self.pages for col in pg.extents)
        res: dict[int, list] = {}
        for i, a in enumerate(anchors):
            nxt = anchors[i + 1] if i + 1 < len(anchors) else None
            nxt_key = (nxt.pno, nxt.col) if nxt else (end[:2] if end else None)
            nxt_y = nxt.y if nxt else (end[2] if end else None)
            rects = []
            k = (a.pno, a.col)
            y = a.y - 2
            while True:
                pg = self.pages[k[0]]
                top, bottom, x0, x1 = pg.extents[k[1]]
                cx0, cx1 = ((x0 - 4, pg.split - 2.5) if k[1] == 0
                            else (pg.split + 2.5, x1 + 3))
                if nxt_key == k:
                    rects.append((k[0], fitz.Rect(cx0, y, cx1, nxt_y - 2)))
                    break
                rects.append((k[0], fitz.Rect(cx0, y, cx1, bottom + 2)))
                if k not in keys or keys.index(k) + 1 >= len(keys):
                    break
                k = keys[keys.index(k) + 1]
                y = self.pages[k[0]].extents[k[1]][0] - 2
                if nxt_key == k and nxt_y <= self.pages[k[0]].extents[k[1]][0] + 1:
                    break  # the next block starts at this column's top
                # with no next block and no end marker (last solution), run on
            res[a.number] = [(p, r) for p, r in rects if r.height > 1]
        return res

    # ── describing ────────────────────────────────────────────────────────
    def preview(self, a: _Anchor) -> str:
        pg = self.pages[a.pno]
        for ln in pg.lines:
            if ln.col == a.col and ln.y0 >= a.y - 2 and QRE.match(ln.text):
                raw = QRE.sub("", ln.text, count=1).strip()
                return (raw[:77] + "…") if len(raw) > 80 else raw
        return ""


# ═══════════════════════════════════════════════════════════════════════════════
# BLOCK → WORD
# ═══════════════════════════════════════════════════════════════════════════════


def _needs_picture(paper: PdfPaper, parts, size: float) -> bool:
    """A block with a photo, a graph or stacked maths is kept as a picture."""
    for pno, r in parts:
        page = paper.doc[pno]
        if any(ib.intersects(r) for _c, ib in paper.pages[pno].images):
            return True
        if any(fitz.Rect(d["rect"]).intersects(r) and d["rect"].width < r.width * 0.9
               for d in page.get_drawings()):
            return True
        ys = sorted(sp["origin"][1] for b in page.get_text("dict")["blocks"]
                    if b["type"] == 0 for l in b["lines"] for sp in l["spans"]
                    if sp["text"].strip() and fitz.Rect(sp["bbox"]).intersects(r)
                    and sp["size"] >= size * 0.85)
        if any(size * 0.3 < b - a < size * 0.85 for a, b in zip(ys, ys[1:])):
            return True
        if any(sp["size"] > size * 1.5 for b in page.get_text("dict")["blocks"]
               if b["type"] == 0 for l in b["lines"] for sp in l["spans"]
               if sp["text"].strip() and fitz.Rect(sp["bbox"]).intersects(r)):
            return True
    return False


def _visual_lines(paper: PdfPaper, parts, size: float):
    """Spans of the block clustered into printed lines by baseline.

    Returns ``[(x0, [run, …])]``; a run is ``dict(text, sup, sub, bold, italic,
    gap)`` — small spans above/below the baseline become super/subscripts.
    """
    spans = []
    for pno, r in parts:
        for b in paper.doc[pno].get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                for sp in l["spans"]:
                    bb = fitz.Rect(sp["bbox"])
                    if sp["text"].strip() and bb.intersects(r) and bb.y0 >= r.y0 - 1:
                        spans.append(sp)
    spans.sort(key=lambda sp: (sp["origin"][1], sp["bbox"][0]))
    lines: list[list] = []  # [baseline, [spans]]
    for sp in spans:
        if sp["size"] < size * 0.85:
            continue
        y = sp["origin"][1]
        for ln in lines:
            if abs(ln[0] - y) <= size * 0.35:
                ln[1].append(sp)
                break
        else:
            lines.append([y, [sp]])
    for sp in spans:  # small spans attach to the nearest line
        if sp["size"] >= size * 0.85:
            continue
        y = sp["origin"][1]
        if lines:
            ln = min(lines, key=lambda l: abs(l[0] - y))
            if abs(ln[0] - y) <= size * 0.9:
                ln[1].append(sp)
                continue
        lines.append([y, [sp]])
    lines.sort(key=lambda l: l[0])
    out = []
    for base, sps in lines:
        sps.sort(key=lambda sp: sp["bbox"][0])
        base_main = max((sp["origin"][1] for sp in sps if sp["size"] >= size * 0.85),
                        default=base)
        runs = []
        last_x1 = None
        for sp in sps:
            t = _span_text(sp)
            if last_x1 is not None and sp["bbox"][0] - last_x1 > size * 0.2 and \
                    not t.startswith(" ") and runs and not runs[-1]["text"].endswith(" "):
                t = " " + t
            small = sp["size"] < size * 0.85
            runs.append(dict(
                text=t,
                sup=small and sp["origin"][1] < base_main - 1,
                sub=small and sp["origin"][1] > base_main + 1,
                bold=bool(sp["flags"] & 16), italic=bool(sp["flags"] & 2),
                gap=last_x1 is not None and sp["bbox"][0] - last_x1 > size * 1.5))
            last_x1 = sp["bbox"][2]
        out.append((sps[0]["bbox"][0], runs))
    return out


def _runs_text(runs) -> str:
    return re.sub(r"[ \t]{2,}", " ", "".join(r["text"] for r in runs)).strip()


def _add_text_block(d, new_no: int, lines, keep_answer: bool = False) -> None:
    """Write a block as editable paragraphs.

    Options ``(n)`` start a paragraph (two on one printed line stay side by
    side), labels start one, other lines continue the previous paragraph.
    """
    paras: list[list] = []  # [kind, runs]
    for i, (x0, runs) in enumerate(lines):
        t = _runs_text(runs)
        if i == 0:
            m = re.match(r"^\s*\*?\s*\d{1,3}\.\s*", t)  # drop the old number
            if m:
                cut, acc = m.end(), 0
                for r in runs:
                    n = len(r["text"])
                    if acc + n <= cut:
                        r["text"] = ""
                        acc += n
                    else:
                        r["text"] = r["text"][max(0, cut - acc):]
                        break
                t = _runs_text(runs)
                if not t:
                    continue
        kind = "opt" if OPT_RE.match(t) else ("label" if LABEL_RE.match(t) else "body")
        if kind == "opt" and len(OPT_RE.findall(t)) >= 2:
            idx = next((k for k, r in enumerate(runs)
                        if k and r["gap"] and OPT_RE.match(r["text"].strip())), None)
            if idx:
                left, right = runs[:idx], runs[idx:]
                right[0]["text"] = right[0]["text"].lstrip()
                tab = dict(text="\t", sup=False, sub=False, bold=False, italic=False, gap=False)
                paras.append(["optrow", left + [tab] + right])
                continue
        if kind in ("opt", "label") or not paras or (i == 0 and keep_answer):
            paras.append([kind, runs])
        else:
            prev = paras[-1]
            if prev[1] and not _runs_text(prev[1]).endswith("-"):
                runs[0]["text"] = " " + runs[0]["text"].lstrip()
            prev[1].extend(runs)

    p = d.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.left_indent = Mm(7)
    p.paragraph_format.first_line_indent = Mm(-7)
    p.paragraph_format.tab_stops.add_tab_stop(Mm(7))
    nr = p.add_run(f"{new_no}.\t")
    nr.bold = True
    nr.font.size = Pt(BODY_PT)
    for k, (kind, runs) in enumerate(paras):
        if k:
            p = d.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.left_indent = Mm(7)
            if kind == "optrow":
                p.paragraph_format.tab_stops.add_tab_stop(Mm(42))
        for r in runs:
            if not r["text"]:
                continue
            text = r["text"] if r["text"] == "\t" else re.sub(r"[ \t]{2,}", " ", r["text"])
            run = p.add_run(text)
            run.font.size = Pt(BODY_PT)
            run.font.superscript = r["sup"]
            run.font.subscript = r["sub"]
            run.bold = r["bold"]
            run.italic = r["italic"]
    d.add_paragraph().paragraph_format.space_after = Pt(2)


def _block_picture(paper: PdfPaper, a: _Anchor, parts, new_no: int) -> tuple[bytes, float]:
    """Render the block's clips (stacked) to PNG with the number rewritten."""
    w = max(r.width for _p, r in parts)
    h = sum(r.height for _p, r in parts)
    tmp = fitz.open()
    page = tmp.new_page(width=w, height=h)
    y = 0.0
    for j, (pno, r) in enumerate(parts):
        page.show_pdf_page(fitz.Rect(0, y, r.width, y + r.height), paper.doc, pno, clip=r)
        if j == 0:
            nb = a.nbox
            dx, dy = -r.x0, y - r.y0
            page.draw_rect(fitz.Rect(nb.x0 + dx - 2, nb.y0 + dy - 1, a.x_stem + dx - 1,
                                     nb.y1 + dy + 1), color=None, fill=(1, 1, 1))
            page.insert_text((nb.x0 + dx, nb.y1 + dy - 2), f"{new_no}.",
                             fontsize=a.size, fontname="hebo")
        y += r.height
    png = page.get_pixmap(dpi=PICTURE_DPI, alpha=False).tobytes("png")
    tmp.close()
    return png, w


def _new_document(paper: PdfPaper):
    """A4, two columns, the source's running header as a picture."""
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(12)
    sec.top_margin, sec.bottom_margin = Mm(12), Mm(12)
    style = d.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(BODY_PT)
    hdr = paper.doc[0]
    hb = fitz.Rect(0, 0, hdr.rect.width, paper.header_h)
    png = hdr.get_pixmap(dpi=PICTURE_DPI, clip=hb, alpha=False).tobytes("png")
    hp = sec.header.paragraphs[0]
    hp.add_run().add_picture(io.BytesIO(png),
                             width=sec.page_width - sec.left_margin - sec.right_margin)
    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:sep"), "1")
    cols.set(qn("w:space"), "500")
    sec._sectPr.append(cols)
    col_w = (sec.page_width - sec.left_margin - sec.right_margin - Emu(Pt(25))) / 2
    return d, col_w


def _emit_block(d, col_w, paper: PdfPaper, a: _Anchor, parts, new_no: int,
                editable: bool, keep_answer: bool = False) -> None:
    if editable and not _needs_picture(paper, parts, a.size):
        _add_text_block(d, new_no, _visual_lines(paper, parts, a.size), keep_answer)
        return
    png, w = _block_picture(paper, a, parts, new_no)
    p = d.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run().add_picture(io.BytesIO(png), width=min(col_w, Emu(Pt(w))))


# ═══════════════════════════════════════════════════════════════════════════════
# PUBLIC API (mirrors paper_core)
# ═══════════════════════════════════════════════════════════════════════════════


def parse_pdf(path: str | Path) -> ParsedPaper:
    """Locate the questions of a PDF → ``ParsedPaper(subjects, has_answer_key)``."""
    paper = PdfPaper(path)
    subject = _implicit_subject(path)
    subjects: OrderedDict[str, list[QuestionInfo]] = OrderedDict()
    subjects[subject] = [QuestionInfo(a.number, subject, paper.preview(a))
                         for a in paper.questions]
    paper.doc.close()
    return ParsedPaper(subjects, bool(paper.answers))


def generate_pdf_paper(source_path, selected_questions: list[QuestionInfo],
                       output_path, editable: bool = True) -> Path:
    """Build a Word question paper from a PDF source. Returns the path written."""
    if not selected_questions:
        raise ValueError("No questions selected.")
    paper = PdfPaper(source_path)
    by_num = {a.number: a for a in paper.questions}
    d, col_w = _new_document(paper)
    number = 0
    for qi in selected_questions:
        a = by_num.get(qi.number)
        if a is None or not paper.q_regions.get(qi.number):
            continue
        number += 1
        _emit_block(d, col_w, paper, a, paper.q_regions[qi.number], number, editable)
    out = _save_with_fallback(d, Path(output_path))
    paper.doc.close()
    return out


def generate_pdf_answer_key(source_path, selected_questions: list[QuestionInfo],
                            output_path, editable: bool = True) -> Path:
    """Answer key + solutions for the selection, numbered like the paper."""
    if not selected_questions:
        raise ValueError("No questions selected.")
    paper = PdfPaper(source_path)
    if not paper.answers:
        paper.doc.close()
        raise ValueError("No answer key was detected in this paper.")
    numbered = []
    for qi in selected_questions:
        if qi.number in paper.q_regions:
            numbered.append((len(numbered) + 1, qi.number))

    d, col_w = _new_document(paper)
    h = d.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = h.add_run("ANSWER KEY")
    hr.bold = True
    hr.underline = True
    hr.font.size = Pt(12)
    for new_no, num in numbered:
        ans = paper.answers.get(num)
        if ans is None:
            continue
        p = d.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.tab_stops.add_tab_stop(Mm(10))
        r = p.add_run(f"{new_no}.\t")
        r.bold = True
        p.add_run(ans)

    sol_by_num = {a.number: a for a in paper.solutions}
    if any(num in sol_by_num for _n, num in numbered):
        d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        for new_no, num in numbered:
            a = sol_by_num.get(num)
            if a is None or not paper.s_regions.get(num):
                continue
            _emit_block(d, col_w, paper, a, paper.s_regions[num], new_no, editable,
                        keep_answer=True)
    out = _save_with_fallback(d, Path(output_path))
    paper.doc.close()
    return out
