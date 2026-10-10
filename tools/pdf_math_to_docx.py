"""STEM exam PDF -> editable Word (.docx) with NATIVE Word equations.

Engines (step 1, recognition)
  mineru  (default) free, runs locally, no key, no upload. MinerU parses the page
          images (OCR mode) into typed blocks: text, inline/display LaTeX, tables,
          figures, each with its position on the page.
  mathpix paid cloud API (MATHPIX_APP_ID / MATHPIX_APP_KEY), best math accuracy.

Then, for both engines
  2. PyMuPDF     : every figure re-cropped from the ORIGINAL PDF at --dpi and
                   inserted at its true printed size
  3. Pandoc      : every $...$ becomes an OMML equation (Insert > Equation object,
                   fully editable); HTML tables keep merged cells
  4. python-docx : page size / margins / column count copied from the PDF, hanging
                   bold question numbers, MCQ options laid out 4-up / 2-up / 1-up

Engine results are cached next to the PDF (<stem>_mineru/ or <stem>_mathpix/), so
re-running with other layout flags takes seconds.

Setup (once)
  pip install pymupdf python-docx requests pypandoc_binary
  MinerU: its own venv, e.g.  python -m venv mineru_env
                              mineru_env\\Scripts\\pip install torch torchvision \\
                                  --index-url https://download.pytorch.org/whl/cu124
                              mineru_env\\Scripts\\pip install mineru
          (models download automatically on the first run, a few GB)

Usage
  python pdf_math_to_docx.py paper.pdf                       # -> paper.docx
  python pdf_math_to_docx.py paper.pdf --pages 1-3 -o test.docx
  python pdf_math_to_docx.py paper.pdf --engine mathpix
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import fitz  # PyMuPDF
import requests
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import house_style  # noqa: E402
from figure_cut import Cut, FigureCutter  # noqa: E402
import paper_format  # noqa: E402


def page_words(page: fitz.Page, clip=None) -> list:
    """The page's words, numbers and option labels in the internal form
    ("Q.12" -> "12.", "A)" -> "(A)") whatever style the paper uses."""
    return paper_format.FORMAT.words(page.get_text("words", clip=clip))

API = "https://api.mathpix.com/v3"
TAB = "⟪TAB⟫"            # placeholder that survives Pandoc -> real <w:tab/>
IMG = "⟪IMG:{}⟫"         # figure placeholder, resolved when rendering
IMG_TOKEN_RE = re.compile("⟪IMG:(\\d+)⟫")

MATH_RE = re.compile(r"\$\$.+?\$\$|\$[^$\n]+?\$", re.S)
MASK_RE = re.compile(r"\$\$.+?\$\$|\$[^$\n]+?\$|<[^>]+>|⟪IMG:\d+⟫", re.S)
MD_IMG_RE = re.compile(r"!\[([^\]]*)\]\((https?://[^)\s]+)\)(\{[^}]*\})?")
QNUM_RE = re.compile(r"^(\s*\*?\s*)(\d{1,3})\.(?=\s)")  # "12." and starred "*12."
OPT_LABEL_RE = re.compile(r"(?:(?<=[\s\0])|^)\(([1-5]|[a-eA-E]|iv|v|i{1,3})\)(?=[\s\0$]|$)")
LABEL_ONLY_RE = re.compile(r"^\s*\(([1-5]|[a-eA-E]|iv|v|i{1,3})\)\s*$")
LOOSE_LABEL_RE = re.compile(r"\(\s+([1-5a-eA-E])\s*\)|\(([1-5a-eA-E])\s+\)")
OPT_FAMILIES = (tuple("12345"), tuple("abcde"), tuple("ABCDE"), ("i", "ii", "iii", "iv", "v"))
DROP_BLOCKS = {"header", "footer", "page_header", "page_footer", "page_number",
               "aside_text", "page_footnote"}
TEXT_LAYER_BLOCKS = {"index"}


@dataclass
class Figure:
    path: Path
    width_pt: float | None
    max_pt: float | None = None
    left_pt: float | None = None   # distance from its column's left edge in the PDF
    centred: bool = False          # printed in the middle of its column
    svg: Path | None = None        # the same crop as vector art (line drawings)

    @property
    def width_in(self) -> float | None:
        if self.width_pt is None:
            return None
        return min(self.width_pt, self.max_pt or 1e9) / 72


@dataclass
class Content:
    lines: list[str]                    # one paragraph (HTML) / one md line each
    figures: dict[int, Figure] = field(default_factory=dict)
    fmt: str = "html"                   # "html" (MinerU) or "markdown" (Mathpix)


# --------------------------------------------------------------------------- #
# Page geometry of the source PDF
# --------------------------------------------------------------------------- #
@dataclass
class Layout:
    width: float   # points
    height: float
    left: float
    right: float
    top: float
    bottom: float
    columns: int
    gutter: float

    @property
    def col_width(self) -> float:
        usable = self.width - self.left - self.right
        return (usable - self.gutter * (self.columns - 1)) / self.columns


def analyse_layout(doc: fitz.Document, columns: str) -> Layout:
    page = doc[0]
    W, H = page.rect.width, page.rect.height
    rects = []
    for p in doc.pages(0, min(doc.page_count, 6)):
        for b in p.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                r = fitz.Rect(ln["bbox"])
                if not r.is_empty and "".join(s["text"] for s in ln["spans"]).strip():
                    rects.append(r)
    if not rects:  # scanned PDF without a text layer
        return Layout(W, H, 36, 36, 36, 36, 2 if columns == "auto" else int(columns), 18)

    def pct(v, q):
        v = sorted(v)
        return v[min(len(v) - 1, max(0, int(q * len(v))))]

    left, right = pct([r.x0 for r in rects], 0.01), W - pct([r.x1 for r in rects], 0.99)
    top, bottom = pct([r.y0 for r in rects], 0.01), H - pct([r.y1 for r in rects], 0.99)
    mid = (left + W - right) / 2
    crossing = sum(1 for r in rects if r.x0 < mid - 6 and r.x1 > mid + 6)
    n = int(columns) if columns != "auto" else (2 if crossing / len(rects) < 0.12 else 1)
    gutter = 18.0
    if n == 2:
        lefts = [r.x1 for r in rects if r.x1 <= mid + 2]
        rights = [r.x0 for r in rects if r.x0 >= mid - 2]
        if lefts and rights:
            gutter = min(48.0, max(10.0, pct(rights, 0.02) - pct(lefts, 0.98)))
    m = lambda v: max(18.0, min(v - 4, 90.0))  # noqa: E731  keep margins printable
    return Layout(W, H, m(left), m(right), m(top), m(bottom), n, gutter)


def crop(doc: fitz.Document, pno: int, rect: fitz.Rect, out: Path, dpi: int) -> Figure:
    page = doc[pno]
    rect = (rect + (-2, -2, 2, 2)) & page.rect
    page.get_pixmap(clip=rect, dpi=dpi).save(out)
    return Figure(out, rect.width)


PROSE_WORDS = {"is", "to", "and", "or", "of", "in", "for", "if", "then", "at", "the", "from",
               "when", "where", "find", "with", "by", "on", "as", "are", "be", "versus", "vs"}


MATH_WORDS = {"sinh", "cosh", "tanh", "arcsin", "arccos", "arctan"}


def _atop_to_matrix(tex: str) -> str:
    r"""{a \atop b} -> \begin{matrix}a\\b\end{matrix} (Pandoc cannot read \atop)."""
    for _ in range(10):
        k = tex.find("\\atop")
        if k < 0:
            break
        depth, a = 0, k
        while a > 0:  # the "{" that opens the group holding \atop
            a -= 1
            depth += {"}": 1, "{": -1}.get(tex[a], 0)
            if depth < 0:
                break
        depth, b = 0, k
        while b < len(tex) - 1:  # its closing "}"
            b += 1
            depth += {"{": 1, "}": -1}.get(tex[b], 0)
            if depth < 0:
                break
        if tex[a] != "{" or tex[b] != "}":
            tex = tex.replace("\\atop", "\\over", 1)  # unbalanced: a plain fraction
            continue
        top, bottom = tex[a + 1:k], tex[k + 5:b]
        tex = (tex[:a] + "\\begin{matrix}" + top.strip() + " \\\\ " + bottom.strip()
               + "\\end{matrix}" + tex[b + 1:])
    return tex


def clean_latex(tex: str) -> str:
    tex = re.sub(r"\s+", " ", tex).strip()
    # old TeX switches Pandoc cannot read ({\bf x}, \tiny); OCR's box for an
    # equilibrium arrow (A + B 口 C + D)
    tex = re.sub(r"\\(?:bf|rm|it|sf|tt|tiny|scriptsize|footnotesize|small|large|Large|LARGE)"
                 r"(?![A-Za-z])\s*", "", tex)
    tex = tex.replace("口", r"\rightleftharpoons ").replace("​", "")
    tex = re.sub(r"\\AA(?![A-Za-z])\s*", "Å", tex)  # ångström: Pandoc cannot read \AA
    tex = _atop_to_matrix(tex)
    # a script with nothing to attach to ("( ^ { C_3H_6 } )" = OCR saw raised text)
    # shows as an empty placeholder box in Word
    for _ in range(3):
        tex = re.sub(r"(^|[{(\[,=]|\\mathrm\s*\{)\s*[\^_]\s*\{(?=\s*(?:\\mathrm\s*\{\s*)?[A-Z])",
                     r"\1{", tex)  # only a raised formula (C_3H_6), never ^{\circ} / ^{2-}
    tex = re.sub(r"\^\s*\{\s*\^\s*\{\s*\\circ\s*\}\s*\}", r"^{\\circ}", tex)  # ^{^{\circ}}
    tex = re.sub(r"(?<=\d) (?=\d)", "", tex)                                  # "1 0" -> "10"
    # OCR spaces out upright text: \mathrm { t h } -> \mathrm{th}, "I V" -> "IV"
    def join_letters(s: str) -> str:  # "A - I V ; t h" -> "A-IV;th", "Matches II" kept
        return re.sub(r"(?<![^\W\d_])([^\W\d_]) (?=[^\W\d_](?![^\W\d_]))|(?<=\W) | (?=\W)",
                      lambda m: m.group(1) or "", s.strip())
    tex = re.sub(r"\\(mathrm|textrm)\s*\{([^{}]*)\}",
                 lambda m: "\\%s{%s}" % (m.group(1), join_letters(m.group(2))), tex)

    # English words that ended up inside math ("t = 0 s to t = 9 s", "\mathrm{is}")
    # OCR-spaced words outside \mathrm: "t a k e n" (4+ single letters) -> "taken"
    tex = re.sub(r"(?<![\\A-Za-z])[A-Za-z](?: [A-Za-z]){3,}(?![A-Za-z])",
                 lambda m: m.group(0).replace(" ", ""), tex)

    def prose(m: re.Match) -> str:
        if m.group(0).startswith("\\"):  # \text{..}, \begin{array}, ...: leave alone
            if not m.group(0).startswith(("\\mathrm", "\\textrm")):
                return m.group(0)
        w = m.group(2) or m.group(3)
        long_word = len(w) >= 4 and w.lower() not in MATH_WORDS
        return f"\\text{{ {w} }}" if w.lower() in PROSE_WORDS or long_word else m.group(0)
    return re.sub(r"\\(?:text|begin|end|operatorname|mathbf|mathit|mathcal|mathbb|boldsymbol"
                  r"|mbox|hbox)\s*\{[^{}]*\}|\\(?:mathrm|textrm)\{(\s*)([A-Za-z]+)\s*\}"
                  r"|(?<![\\A-Za-z])([A-Za-z]{2,})(?![A-Za-z])", prose, tex)


ITEM_RE = re.compile(r"(?:(?<=\s)|^)(VI|IV|V|I{1,3}|[A-E])\.\s")
ITEM_FAMILIES = (("I", "II", "III", "IV", "V", "VI"), ("A", "B", "C", "D", "E"))
# "Statement-I:", "Statement I.", "Assertion (A):", "Assertion A:", "Reason R:", "Choose the ..."
STATEMENT_RE = re.compile(r"\s+(?=(?:<b>)?(?:Statement|Assertion|Reason)\s*[-\u2013]?\s*"
                          r"(?:(?:[IVX]+|\d+|\(?[A-Z]\)?)?\s*:|(?:[IVX]+|\d+)\.\s)"
                          r")|\s+(?=(?:<b>)?Choose the\b)")


# --------------------------------------------------------------------------- #
# 1a. MinerU (local)
# --------------------------------------------------------------------------- #
def find_mineru(explicit: str | None) -> list[str]:
    """Command that runs mineru-kit.

    MinerU installed in THIS interpreter (the portable Paper Maker engine) runs as
    a module: pip's mineru-kit.exe launcher stores an absolute python path, which
    breaks once the folder is copied to another PC.
    """
    cands = [explicit, os.environ.get("MINERU_KIT")]
    if not any(cands):
        try:
            import importlib.util
            if importlib.util.find_spec("mineru.kit.main"):
                return [sys.executable, "-E", "-s", "-X", "utf8", "-c",
                        "import sys; from mineru.kit.main import main; sys.exit(main())"]
        except ImportError:
            pass
    cands += [shutil.which("mineru-kit"), str(Path(sys.executable).with_name("mineru-kit.exe")),
              str(Path(__file__).resolve().parents[2] / "mineru_env" / "Scripts" / "mineru-kit.exe")]
    for c in cands:
        if c and Path(c).exists():
            return [c]
    sys.exit("mineru-kit not found: install MinerU (see the header of this script) or pass --mineru")


def parse_pages(spec: str, count: int) -> list[int]:
    """'1-3,5' -> [0, 1, 2, 4] (0-based); 'all' -> every page."""
    if spec.strip().lower() == "all":
        return list(range(count))
    out = []
    for part in spec.split(","):
        a, _, b = part.strip().partition("-")
        out += range(int(a) - 1, int(b or a))
    return [p for p in dict.fromkeys(out) if 0 <= p < count]


def mineru_convert(pdf: Path, cache: Path, exe: list[str], pages: list[int], tier: str) -> None:
    if (cache / "middle_json.json").exists():
        print(f"[mineru] using cached results in {cache}")
        return
    cache.mkdir(exist_ok=True)
    src = cache / "pages.pdf"  # exactly the selected pages, so page_idx maps 1:1
    with fitz.open(pdf) as d, fitz.open() as sub:
        for p in pages:
            sub.insert_pdf(d, from_page=p, to_page=p)
        sub.save(src)
    print(f"[mineru] parsing {len(pages)} page(s) locally (tier={tier}, OCR mode); "
          "the first run downloads models")
    t = time.time()
    # OCR mode: in text-layer mode MinerU silently drops formulas typeset in
    # math fonts (e.g. "measuring __ in length"); OCR mode recovers them.
    subprocess.run([*exe, "parse", str(src), "-o", str(cache / "out.zip"), "-f", "zip",
                    "--tier", tier, "--ocr-mode", "ocr", "-p", "all"], check=True)
    with zipfile.ZipFile(cache / "out.zip") as z:
        z.extractall(cache)
    print(f"[mineru] done in {time.time() - t:.0f}s")


def _spans_to_html(content) -> str:
    """A block's text as HTML with $maths$, numbers and labels in the internal form."""
    return paper_format.FORMAT.line(_spans_html(content))


def _spans_html(content) -> str:
    if isinstance(content, str):
        return html.escape(content)
    out = []
    for s in content:
        t, c = s.get("type"), s.get("content", "")
        if t in ("equation_inline", "inline_equation"):
            out.append(f" ${html.escape(clean_latex(c))}$ ")
        elif t in ("equation", "interline_equation"):
            out.append(f" $${html.escape(clean_latex(c))}$$ ")
        elif isinstance(c, list):
            out.append(_spans_html(c))
        else:
            c = re.sub("[①-⑤]", lambda m: f" ({ord(m.group()) - 0x245F}) ", c)  # ① -> (1)
            c = c.replace("​", "")
            txt = html.escape(LOOSE_LABEL_RE.sub(lambda m: f"({m.group(1) or m.group(2)})", c))
            if "bold" in (s.get("styles") or []) and not QNUM_RE.match(txt + " "):
                txt = f"<b>{txt}</b>"
            out.append(txt)
    return re.sub(r"[ \t]{2,}", " ", "".join(out)).strip()


def reading_key(bbox, mid: float, cuts: list[float], tol: float) -> tuple:
    """(band, column, y) - the order a person reads a 2-column exam page.

    A full-width element (subject heading "CHEMISTRY", a cover line) cuts the
    page into bands: everything above it is read first (left column, then right),
    then the element itself, then the band below it.
    """
    x0, y0, x1, _ = bbox
    band = sum(1 for c in cuts if c <= y0 + tol * 0.1)
    col = 0 if (x0 < mid - tol and x1 > mid + tol) else int((x0 + x1) / 2 >= mid)
    return band, col, y0


def option_label_words(page: fitz.Page) -> list[tuple]:
    """Text-layer words that are option labels: [(rect, "(2)")]."""
    return [(fitz.Rect(w[:4]), w[4]) for w in page_words(page)
            if re.fullmatch(r"\(\s*(?:[1-5]|[a-eA-E])\s*\)", w[4])]


def label_left_of(labels: list[tuple], fig: fitz.Rect, gap: float = 40) -> str | None:
    """The option label printed just left of a figure, on the same row."""
    hit = label_hit(labels, fig, gap)
    return hit[1] if hit else None


def label_hit(labels: list[tuple], fig: fitz.Rect, gap: float = 40) -> tuple | None:
    best = None
    for r, t in labels:
        cy = (r.y0 + r.y1) / 2
        if fig.y0 - 4 <= cy <= fig.y1 + 4 and r.x1 <= fig.x0 + 4 and fig.x0 - r.x1 < gap:
            if best is None or r.x1 > best[0].x1:
                best = (r, t)
    return best


def column_bounds(doc: fitz.Document) -> list[tuple[float, float]]:
    """[(left, right)] of the text columns, from the lines of the first pages."""
    W = doc[0].rect.width
    mid = W / 2
    rects = [fitz.Rect(ln["bbox"]) for p in doc.pages(0, min(doc.page_count, 6))
             for b in p.get_text("dict")["blocks"] for ln in b.get("lines", [])
             if "".join(s["text"] for s in ln["spans"]).strip()]
    if not rects:
        return [(36.0, mid - 9), (mid + 9, W - 36.0)]

    def pct(v, q):
        v = sorted(v)
        return v[min(len(v) - 1, max(0, int(q * len(v))))]

    crossing = sum(1 for r in rects if r.x0 < mid - 6 and r.x1 > mid + 6)
    if crossing / len(rects) >= 0.12:  # one column
        return [(pct([r.x0 for r in rects], 0.02), pct([r.x1 for r in rects], 0.98))]
    a = [r for r in rects if r.x1 <= mid + 2] or rects
    b = [r for r in rects if r.x0 >= mid - 2] or rects
    return [(pct([r.x0 for r in a], 0.02), pct([r.x1 for r in a], 0.98)),
            (pct([r.x0 for r in b], 0.02), pct([r.x1 for r in b], 0.98))]


def place_in_column(fig: Figure, r: fitz.Rect, cols: list[tuple[float, float]]) -> None:
    """Remember where the figure stood: centred in its column, or how far in."""
    x0, x1 = min(cols, key=lambda c: abs((c[0] + c[1]) / 2 - (r.x0 + r.x1) / 2))
    fig.left_pt = max(0.0, r.x0 - x0)
    fig.centred = abs((r.x0 + r.x1) / 2 - (x0 + x1) / 2) <= 0.07 * (x1 - x0)


def split_at_labels(page: fitz.Page, r: fitz.Rect, labels: list[tuple]) -> list[tuple]:
    """[(rect, label)] - one picture box that holds several option figures stacked
    ("(A)" and "(C)" both printed left of it) is cut between the labels, at the
    emptiest height between them."""
    rows = sorted(((lr, t) for lr, t in labels
                   if r.y0 - 4 <= (lr.y0 + lr.y1) / 2 <= r.y1 + 4 and lr.x1 <= r.x0 + 4
                   and r.x0 - lr.x1 < 40), key=lambda x: x[0].y0)
    rows = [x for i, x in enumerate(rows) if i == 0 or x[0].y0 - rows[i - 1][0].y0 > 6]
    if len(rows) < 2:
        return [(r, rows[0][1] if rows else None)]
    pm = page.get_pixmap(clip=r, dpi=72, colorspace=fitz.csGRAY)
    ink = [sum(1 for v in pm.samples[y * pm.stride:y * pm.stride + pm.width] if v < 200)
           for y in range(pm.height)]
    cuts = []
    for (a, _), (b, _) in zip(rows, rows[1:]):
        lo = max(0, int((a.y0 + a.y1) / 2 - r.y0))
        hi = min(pm.height - 1, int((b.y0 + b.y1) / 2 - r.y0))
        if hi <= lo:
            return [(r, rows[0][1])]
        least = min(ink[lo:hi + 1])
        blank = [y for y in range(lo, hi + 1) if ink[y] == least]
        cuts.append(r.y0 + blank[len(blank) // 2])
    edges = [r.y0] + cuts + [r.y1]
    return [(fitz.Rect(r.x0, y0, r.x1, y1), t) for (y0, y1), (_, t) in zip(zip(edges, edges[1:]), rows)]


def table_tail_options(body: str) -> tuple[str, list[str]]:
    """(table, options) - the "(A) I-P; II-Q ..." rows OCR pulled into the end of a
    Match table are the question's options, not table rows."""
    rows = re.findall(r"<tr[^>]*>.*?</tr>", body, re.S)
    tail: list[str] = []
    while len(rows) - len(tail) > 1:
        cells = [c.strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", rows[-1 - len(tail)], re.S)]
        cells = [c for c in cells if re.sub(r"<[^>]+>|\s", "", c)]
        if len(cells) != 1 or not OPT_LABEL_RE.match(_masked(cells[0])):
            break
        tail.insert(0, cells[0])
    if not tail or _label_index(tail[0])[1] != 0:
        return body, []
    spaced = [re.sub(r"(?<=[^\s(])(\((?:[2-5]|[B-Eb-e])\)\s)", r" \1", t) for t in tail]
    tail = [o for t in spaced for o in (_split_options(t) or [t])]
    for row in rows[len(rows) - len(tail):]:
        body = body.replace(row, "", 1)
    return body, tail


def rows_left_to_right(items: list, box, key) -> list:
    """Items sorted top-down, with side-by-side items (an option grid "(1) [fig]
    (2) [fig]") read left to right: same band and column, overlapping vertically."""
    out: list = []
    row: list = []

    def flush():
        out.extend(sorted(row, key=lambda it: box(it).x0))
        row.clear()

    for it in items:
        if row:
            a, b = box(row[0]), box(it)
            overlap = min(a.y1, b.y1) - max(a.y0, b.y0)
            same = key(a)[:2] == key(b)[:2] and overlap > 0.5 * min(a.height, b.height)
            if not same:
                flush()
        row.append(it)
    flush()
    return out


def spanning_cuts(bboxes, mid: float, tol: float) -> list[float]:
    """Tops of the full-width elements: they cross the column gutter AND nothing
    stands beside them (a long left-column line that overflows the middle has
    right-column text on the same height)."""
    wide = [b for b in bboxes if b[0] < mid - tol and b[2] > mid + tol]
    narrow = [b for b in bboxes if not (b[0] < mid - tol and b[2] > mid + tol)]
    return sorted(w[1] for w in wide
                  if not any(min(w[3], n[3]) - max(w[1], n[1]) > 0.3 * (w[3] - w[1])
                             for n in narrow))


def text_segments(page: fitz.Page) -> list[list]:
    """The page's words grouped into (band, column) segments, in reading order."""
    W = page.rect.width
    mid, tol = W / 2, 15
    words = page_words(page)
    boxes: dict[tuple, fitz.Rect] = {}  # a line's box without its trailing spaces
    for w in words:
        boxes[w[5], w[6]] = boxes.get((w[5], w[6]), fitz.Rect(w[:4])) | fitz.Rect(w[:4])
    lines = list(boxes.values())
    cuts = spanning_cuts(lines, mid, tol)
    wide = [r for r in lines if r.x0 < mid - tol and r.x1 > mid + tol and r.y0 in cuts]
    segs: dict[tuple, list] = {}
    for w in words:
        c = fitz.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2)
        band, col, _ = reading_key(w[:4], mid, cuts, tol)
        if any(r.contains(c) for r in wide):
            col = 0  # a word of a full-width line stays with its line
        segs.setdefault((band, col), []).append(w)
    return [segs[k] for k in sorted(segs)]


def text_layer_rows(page: fitz.Page, rect: fitz.Rect) -> list[str]:
    """Rows of real text inside rect (2-column aware), straight from the PDF text layer."""
    words = page_words(page, clip=rect)
    if not words:
        return []
    mid = (rect.x0 + rect.x1) / 2
    two_col = not any(w[0] < mid - 4 and w[2] > mid + 4 for w in words)
    rows: dict[tuple[int, int], list] = {}
    for w in words:
        key = (int(two_col and w[0] >= mid), round((w[1] + w[3]) / 2 / 4))
        rows.setdefault(key, []).append(w)
    return [" ".join(w[4] for w in sorted(ws, key=lambda w: w[0]))
            for _, ws in sorted(rows.items())]


try:  # Paper Maker's map for Symbol-font glyphs in PDF text layers (private-use code points)
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from pdf_core import _SYMBOL_MAP
except Exception:  # pragma: no cover - the converter also works standalone
    _SYMBOL_MAP = {}


def missed_text(page: fitz.Page, blocks: list[dict], key) -> list[tuple]:
    """Text-layer lines that no MinerU block covers (its layout model skipped them).

    Returns [(reading-order key, lines, area)] for runs with real text; isolated
    labels, axis letters and page furniture are ignored. ``key(rect)`` gives the
    reading-order key of a rectangle in page points (same order as the blocks).
    """
    W, H = page.rect.width, page.rect.height
    # a block's parts (table/figure captions, footnotes) can lie outside its own box
    parts = [p for b in blocks for p in (b.get("content") if isinstance(b.get("content"), list)
                                         else []) if isinstance(p, dict) and p.get("bbox")]
    boxes = [fitz.Rect(b["bbox"][0] * W - 3, b["bbox"][1] * H - 3,
                       b["bbox"][2] * W + 3, b["bbox"][3] * H + 3) for b in blocks + parts]
    missed = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            r = fitz.Rect(ln["bbox"])
            t = paper_format.FORMAT.line("".join(s["text"] for s in ln["spans"]).strip())
            if not t or r.y0 < H * 0.06 or r.y1 > H * 0.96:
                continue
            c = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
            if not any(bx.contains(c) for bx in boxes):
                missed.append((key(r)[:2], r, t.translate(_SYMBOL_MAP)))
    missed.sort(key=lambda m: (m[0], m[1].y0, m[1].x0))
    runs, cur = [], []
    for m in missed:
        if cur and (m[0] != cur[-1][0] or m[1].y0 - cur[-1][1].y1 > m[1].height * 0.8):
            runs.append(cur); cur = []
        cur.append(m)
    if cur:
        runs.append(cur)
    out = []
    for run in runs:
        if len(re.sub(r"[^A-Za-z0-9]", "", " ".join(m[2] for m in run))) < 12:
            continue
        lines = []
        prev = None
        for col, r, t in run:  # an indented line without a label continues the previous
            beside = prev is not None and abs(r.y0 - prev.y0) < 3 and r.x0 > prev.x1 + 30 \
                and len(t) <= 30 and len(lines[-1]) <= 30
            if lines and beside and not OPT_LABEL_RE.match(t) and not OPT_LABEL_RE.match(lines[-1]):
                lines[-1] += "\0" + t  # "Figure (a)    Figure (b)": one tab cell each
            elif lines and not OPT_LABEL_RE.match(t) and r.x0 > run[0][1].x0 + 8:
                lines[-1] += " " + t
            else:
                lines.append(t)
            prev = r
        area = fitz.Rect(run[0][1])
        for m in run:
            area |= m[1]
        out.append((key(run[0][1]), [html.escape(x).replace("\0", TAB) for x in lines], area))
    return out


def fix_glued_words(lines: list[str], doc: fitz.Document) -> list[str]:
    """'atmidpoint' -> 'at midpoint', 'accelerationtime' -> 'acceleration-time'.

    OCR glues the last word of a wrapped line to the first of the next; the PDF's
    own vocabulary says where to split (and whether the source hyphenated it).
    """
    src = " ".join(p.get_text() for p in doc)
    vocab = {w.lower() for w in re.findall(r"[A-Za-z]+", src)}
    hyphen = {m.group(1).lower() for m in re.finditer(r"([A-Za-z]+)-\s*\n", src)}
    if len(vocab) < 50:
        return lines
    fixed = 0

    def split(m: re.Match) -> str:
        nonlocal fixed
        w = m.group(0)
        if w.lower() in vocab:
            return w
        for i in range(len(w) - 2, 1, -1):  # prefer the longer first word
            a, b = w[:i], w[i:]
            if a.lower() in vocab and b.lower() in vocab and len(b) >= 2:
                fixed += 1
                return f"{a}-{b}" if a.lower() in hyphen else f"{a} {b}"
        return w

    def text_only(s: str) -> str:  # leave maths, tags and figure tokens alone
        parts = MASK_RE.split(s)
        keep = MASK_RE.findall(s)
        parts = [re.sub(r"[A-Za-z]{6,}", split, p) for p in parts]
        return "".join(x for pair in zip(parts, keep + [""]) for x in pair)

    lines = [text_only(ln) for ln in lines]
    if fixed:
        print(f"[repair] split {fixed} word(s) OCR had glued together")
    return lines


ROMAN = ("I", "II", "III", "IV", "V", "VI", "VII", "VIII")
LABEL_CELL_RE = re.compile(r"^\(\s*([A-Za-z0-9]{1,4})\s*\)?$")  # "(II)", OCR "(IVD", "(1)"


def _table_rows(body: str) -> list[list[str]]:
    return [re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)
            for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", body, re.S)]


def repair_tables(lines: list[str]) -> list[str]:
    """Match-the-column tables: glue halves cut at a column/page end, fix label OCR.

    * MinerU emits a table that runs into the next column as two tables; the second
      one's first row (first cell empty) continues the first one's last row.
    * Label columns "(A)..(D)" / "(I)..(IV)" are renumbered in order: OCR reads
      "(I)" as "(1)", "(IV)" as "(IVD", and uses full-width brackets.
    """
    out: list[str] = []
    for ln in lines:
        ln = ln.replace("（", "(").replace("）", ")") if ln.startswith("<table") else ln
        prev = out[-1] if out else ""
        if ln.startswith("<table") and prev.startswith("<table"):
            rows = _table_rows(ln)
            first = rows[0] if rows else None
            if first and not re.sub(r"<[^>]+>|\s", "", first[0]):
                # cell-wise continuation of the previous table's last row
                last = re.findall(r"<tr[^>]*>(?:(?!<tr).)*?</tr>", prev, re.S)[-1]
                cells = re.findall(r"(<td[^>]*>)(.*?)(</td>)", last, re.S)
                if len(cells) == len(first):
                    joined = "<tr>" + "".join(f"{a}{b.strip()} {c_.strip()}{c}"
                                              for (a, b, c), c_ in zip(cells, first)) + "</tr>"
                    prev = prev[:prev.rindex(last)] + joined + prev[prev.rindex(last) + len(last):]
                    rest = re.findall(r"<tr[^>]*>.*?</tr>", ln, re.S)[1:]
                    out[-1] = prev.replace("</table>", "".join(rest) + "</table>")
                    print("[repair] a table cut at a column end was joined")
                    continue
        if ln.startswith("<table"):
            ln = _renumber_labels(ln)
        out.append(ln)
    return out


def _renumber_labels(body: str) -> str:
    rows = _table_rows(body)
    width = max((len(r) for r in rows), default=0)
    for col in range(width):
        cells = [(i, r[col]) for i, r in enumerate(rows) if col < len(r)]
        labs = [(i, LABEL_CELL_RE.match(re.sub(r"<[^>]+>", "", c).strip())) for i, c in cells]
        labs = [(i, m.group(1)) for i, m in labs if m]
        if len(labs) < 3:
            continue
        vals = [v for _, v in labs]
        if sum(v in ROMAN for v in vals) >= len(vals) / 2:
            seq = ROMAN
        elif sum(len(v) == 1 and v.isalpha() and v.isupper() for v in vals) >= len(vals) / 2:
            seq = tuple("ABCDEFGH")
        else:
            continue
        if len(labs) > len(seq):
            continue
        fixes = {i: f"({seq[k]})" for k, (i, _) in enumerate(labs)}
        for i, new in fixes.items():
            rows[i][col] = new
    # write the cell texts back in order
    it = iter([c for r in rows for c in r])
    return re.sub(r"(<td[^>]*>)(.*?)(</td>)", lambda m: m.group(1) + next(it) + m.group(3), body,
                  flags=re.S)


def question_table(body: str) -> list[str] | None:
    """A borderless "stem | (1) a | (2) b" layout misread as a table -> stem + options.

    Real tables (Match List-I/List-II, data) never start with a question number.
    """
    cells = [re.sub(r"<[^>]+>", " ", c).strip()
             for c in re.findall(r"<td[^>]*>(.*?)</td>", body, re.S)]
    cells = [re.sub(r"\s+", " ", c) for c in cells if c.strip()]
    if not cells or not QNUM_RE.match(cells[0] + " "):
        return None
    stem = [c for c in cells if not OPT_LABEL_RE.match(_masked(c))]
    opts = [c for c in cells if OPT_LABEL_RE.match(_masked(c))]
    if len(opts) < 2:
        return None
    print(f"[repair] question {cells[0].split('.')[0]} was read as a table: rebuilt as text")
    return [" ".join(stem)] + opts


def join_pictures(upper: Figure, lower: Figure) -> bool:
    """Stack ``lower`` under ``upper`` in upper's file (a table printed in two parts
    is one table again). False when the two do not have the same width."""
    try:
        from PIL import Image
    except ImportError:
        return False
    if not upper.width_pt or not lower.width_pt \
            or abs(upper.width_pt - lower.width_pt) > 0.08 * upper.width_pt:
        return False
    with Image.open(upper.path) as a, Image.open(lower.path) as b:
        dpi = a.info.get("dpi") or (300, 300)  # Pandoc sizes a picture by its dpi
        a, b = a.convert("RGB"), b.convert("RGB")

        def ink_box(im) -> tuple[int, int, int, int]:  # the table itself, without margins
            box = im.convert("L").point(lambda v: 255 if v < 200 else 0).getbbox()
            return box or (0, 0, im.width, im.height)

        (ax0, _ay0, ax1, ay1), (bx0, by0, bx1, _by1) = ink_box(a), ink_box(b)
        pad = ax0
        a = a.crop((0, 0, a.width, ay1))                 # no blank band between the parts
        b = b.crop((bx0, by0, bx1, b.height))            # left and right edges line up
        width = ax1 - ax0
        b = b.resize((width, max(1, round(b.height * width / b.width))))
        both = Image.new("RGB", (a.width, a.height + b.height), "white")
        both.paste(a, (0, 0))
        both.paste(b, (pad, a.height))
    both.save(upper.path, dpi=dpi)
    upper.svg = None
    return True



def _row_caption(lines: list[str], row: dict, k: int, text: str) -> None:

    """Write `text` under figure k of a side-by-side row (one tab cell per figure)."""
    if row["caps"] is None:
        lines.append("")
        row["caps"] = len(lines) - 1
        row["cells"] = [""] * len(row["rects"])
    row["cells"] += [""] * (len(row["rects"]) - len(row["cells"]))
    row["cells"][k] = (row["cells"][k] + " " + text).strip()
    lines[row["caps"]] = TAB.join(row["cells"])


def mineru_content(cache: Path, doc: fitz.Document, dpi: int, pages: list[int],
                   vector: bool = False) -> Content:
    m = json.loads((cache / "middle_json.json").read_text(encoding="utf-8"))
    fig_dir = cache / "figures"
    fig_dir.mkdir(exist_ok=True)
    c = Content([])
    cols = column_bounds(doc)
    cutter = FigureCutter(doc)
    moved = []  # how far each refined crop is from the layout model's box
    table_pics: set[int] = set()  # figures that are pictures of a table

    def figure(pno: int, bbox, cut: Cut | None = None) -> str:
        page = doc[pno]
        W, H = page.rect.width, page.rect.height
        r = bbox if isinstance(bbox, fitz.Rect) else \
            fitz.Rect(bbox[0] * W, bbox[1] * H, bbox[2] * W, bbox[3] * H)
        k = len(c.figures) + 1
        if cut is None:  # a table kept as a picture: its own box, a little air around
            cut = Cut((r + (-2, -2, 2, 2)) & page.rect)
        out = fig_dir / f"fig{k:03d}.png"
        svg = fig_dir / f"fig{k:03d}.svg" if vector and cut.vector else None
        cutter.render(pno, cut, out, dpi, svg)
        svg = svg if svg is not None and svg.exists() else None  # None: not exact as vector
        c.figures[k] = Figure(out, cut.rect.width, svg=svg)
        place_in_column(c.figures[k], cut.rect, cols)
        return IMG.format(k)

    for page in m["pages"]:
        pno = pages[page["page_idx"]]
        W = doc[pno].rect.width
        page_start = len(c.lines)
        H = doc[pno].rect.height
        rect = lambda bb: fitz.Rect(bb[0] * W, bb[1] * H, bb[2] * W, bb[3] * H)  # noqa: E731
        blocks = []
        for b in page["blocks"]:
            if b["type"] in DROP_BLOCKS:
                continue
            parts = b.get("content") if b["type"] in ("table", "image") else None
            if isinstance(parts, list):
                # a caption MinerU attached from the other column ("178. Match ..." on
                # the table of 175) is a block of its own, read where it stands
                col = int((b["bbox"][0] + b["bbox"][2]) / 2 >= 0.5)
                stray = [p for p in parts if isinstance(p, dict) and p.get("bbox")
                         and not p.get("type", "").endswith("_body")
                         and int((p["bbox"][0] + p["bbox"][2]) / 2 >= 0.5) != col]
                if stray:
                    b = dict(b, content=[p for p in parts if p not in stray])
                    blocks += [{"type": "text", "bbox": p["bbox"], "content": p.get("content", "")}
                               for p in stray]
            blocks.append(b)
        cuts = spanning_cuts([rect(b["bbox"]) for b in blocks], W / 2, W * 0.025)
        key = lambda r: reading_key(r, W / 2, cuts, W * 0.025)  # noqa: E731
        # MinerU can OCR a figure a second time as a text block of garbage over the
        # same area; only a question number printed left of the figure is real
        images = [rect(b["bbox"]) for b in blocks if b["type"] == "image"]
        label_words = option_label_words(doc[pno])
        fig_number: dict[int, str] = {}
        kept = []
        for b in blocks:
            r = rect(b["bbox"])
            lone = "" if b["type"] in ("image", "table") else \
                re.sub(r"<[^>]+>", "", _spans_to_html(b.get("content", ""))).strip()
            if lone and LABEL_ONLY_RE.match(lone) and any(
                    label_left_of([(r, "")], im) is not None for im in images):
                # "(2)" beside an option figure: the figure carries its label (a
                # scanned page has no text layer to read it from again)
                if not any((r & lr).get_area() > 0 for lr, _ in label_words):
                    label_words.append((r, lone))
                continue
            twin = next((im for im in images if b["type"] not in ("image", "table")
                         and (r & im).get_area() > 0.7 * max(r.get_area(), im.get_area())), None)
            if twin is not None:
                s = re.sub(r"<[^>]+>", " ", _spans_to_html(b.get("content", "")))
                n = re.search(r"(?<![\d.])(\d{1,3})\.(?!\d)", s)
                if n and r.x0 < twin.x0 - 0.02 * W:
                    fig_number[id(twin)] = n.group(1)
                    for bb_ in blocks:  # the figure block carries the number
                        if bb_["type"] == "image" and rect(bb_["bbox"]) == twin:
                            fig_number[id(bb_)] = n.group(1)
                continue
            kept.append(b)
        pics = sorted((b for b in kept if b["type"] == "image"), key=lambda b: b["bbox"][1])
        for lo in pics:  # one option figure cut into two stacked boxes: same label
            rl = rect(lo["bbox"])
            for up in pics:
                ru = rect(up["bbox"])
                hit = label_hit(label_words, ru)
                share = lambda a, b: (a & b).get_area() / max(a.get_area(), 1e-6)  # noqa: E731
                # two boxes on ONE embedded picture are one figure, labelled or not
                one = up is not lo and up in kept and ru.y0 <= rl.y0 and any(
                    share(ru, i) > 0.5 and share(rl, i) > 0.5 for i in cutter.images(pno))
                if not one and (
                        up is lo or up not in kept or hit is None or abs(rl.y0 - ru.y1) > 6
                        or min(ru.x1, rl.x1) - max(ru.x0, rl.x0) < 0.6 * min(ru.width, rl.width)
                        or label_hit(label_words, rl) not in (None, hit)):
                    continue
                box = [min(up["bbox"][0], lo["bbox"][0]), up["bbox"][1],
                       max(up["bbox"][2], lo["bbox"][2]), max(up["bbox"][3], lo["bbox"][3])]
                for part in [up] + [q for q in (up.get("content") or [])
                                    if isinstance(q, dict) and q.get("type") == "image_body"]:
                    part["bbox"] = box
                kept.remove(lo)
                break
        blocks = rows_left_to_right(sorted(kept, key=lambda b: key(rect(b["bbox"]))),
                                    lambda b: rect(b["bbox"]), key)
        bodies = lambda b: [q for q in (b.get("content") or []) if isinstance(q, dict)  # noqa: E731
                            and q.get("type") == "image_body"] or [b]
        seeds = [rect(q["bbox"]) for b in kept if b["type"] == "image" for q in bodies(b)]
        # what is never part of a figure: text and table blocks, option labels, captions
        claimed = [rect(b["bbox"]) for b in kept if b["type"] != "image"]
        claimed += [lr for lr, _ in label_words]
        captions = [rect(q["bbox"]) for b in kept if b["type"] == "image"
                    for q in (b.get("content") or [])
                    if isinstance(q, dict) and q.get("bbox") and q.get("type") != "image_body"]
        claimed += captions
        # foreign as whole areas: tables (borders and all), captions printed as text
        solid = captions + [rect(b["bbox"]) for b in kept if b["type"] == "table"]
        used: list[fitz.Rect] = []
        keyed: list[list] = []  # [reading key, first line index, block rect]
        # figures printed side by side without labels ("Figure (a)  Figure (b)") share
        # one row: {"rects", "line" (index in c.lines), "caps" (line index or None)}
        fig_row: dict | None = None
        for b in blocks:
            t = b["type"]
            bb = b["bbox"]
            keyed.append([key(rect(bb)), len(c.lines), rect(bb)])
            if t == "image":
                parts = b.get("content") or []
                body = [p for p in parts if p.get("type") == "image_body"] or [b]
                lead = f"{fig_number[id(b)]}. " if id(b) in fig_number else ""
                # an option label MinerU filed as the picture's caption ("(1)" left of
                # an option figure) goes in front of the picture, not after it
                labels = [p for p in parts if p.get("type") != "image_body" and p.get("bbox")
                          and LABEL_ONLY_RE.match(re.sub(r"<[^>]+>", "",
                                                          _spans_to_html(p.get("content", ""))))
                          and p["bbox"][0] < min(q["bbox"][0] for q in body) + 0.005]
                caps = [_spans_to_html(p["content"]) for p in parts
                        if p.get("type") != "image_body" and p.get("content") and p not in labels]
                for p in body:
                    seed = rect(p["bbox"])
                    cut = cutter.refine(pno, seed, [o for o in seeds if o != seed], claimed,
                                        solid)
                    if any((cut.rect & u).get_area() > 0.8 * min(cut.rect.get_area(), u.get_area())
                           for u in used):
                        continue  # a second box on a figure that is already placed
                    used.append(cut.rect)
                    moved.append(cut.moved)
                    # the PDF's own label beside the picture beats OCR's guess
                    for r, lab in split_at_labels(doc[pno], cut.rect, label_words):
                        part = Cut(r, cut.foreign, cut.wipes, cut.vector)
                        if lab is None and labels:
                            lab = re.sub(r"<[^>]+>", "",
                                         _spans_to_html(labels[0]["content"])).strip()
                        beside = fig_row is not None and not lab and not lead and all(
                            key(r)[:2] == key(o)[:2] for o in fig_row["rects"]) and any(
                            min(r.y1, o.y1) - max(r.y0, o.y0) > 0.5 * min(r.height, o.height)
                            for o in fig_row["rects"])
                        if beside:
                            c.lines[fig_row["line"]] += TAB + figure(pno, r, part)
                            fig_row["rects"].append(r)
                        else:
                            c.lines.append(lead + (f"{lab} " if lab else "")
                                           + figure(pno, r, part))
                            fig_row = None if lab or lead else {
                                "rects": [r], "line": len(c.lines) - 1, "caps": None}
                        lead = ""
                short = fig_row is not None and caps and all(len(x) <= 40 for x in caps)
                if short and len(caps) == len(fig_row["rects"]) > 1:
                    for k, cap in enumerate(caps):  # "Figure(a)", "Figure(b)": one each
                        _row_caption(c.lines, fig_row, k, cap)
                elif short and len(caps) == 1:
                    _row_caption(c.lines, fig_row, len(fig_row["rects"]) - 1, caps[0])
                else:
                    c.lines += caps  # captions / footnotes
            elif t == "table":
                for p in b.get("content") or []:
                    body = p.get("content")
                    if p.get("type") == "table_body" and isinstance(body, str) and "<table" in body:
                        body = re.sub(r"<eq>(.*?)</eq>",
                                      lambda e: f" ${html.escape(clean_latex(html.unescape(e.group(1))))}$ ",
                                      body, flags=re.S)
                        body, tail = table_tail_options(body)
                        if "<img" in body:
                            # drawings inside the cells (structures, graphs): the grid OCR
                            # rebuilds around them is unreliable - keep the table's exact look
                            r = rect(p["bbox"])
                            first = re.sub(r"<[^>]+>", "", tail[0]).split()[0] if tail else None
                            tops = [w[1] for w in page_words(doc[pno], clip=r)
                                    if w[4] == first and w[0] < r.x0 + 0.25 * r.width]
                            if tops:  # the options pulled out below are not part of it
                                r.y1 = max(r.y0 + 10, max(tops) - 1)
                            air = (r + (-2, -2, 2, 2)) & doc[pno].rect
                            # the stem line just above / below is not part of the table
                            outside = [w for w in cutter.words(pno) if not r.contains(
                                fitz.Point((w.x0 + w.x1) / 2, (w.y0 + w.y1) / 2))]
                            token = figure(pno, r, Cut(air, outside))
                            k = int(IMG_TOKEN_RE.fullmatch(token).group(1))
                            prev = IMG_TOKEN_RE.fullmatch(c.lines[-1].strip()) if c.lines else None
                            if prev and int(prev.group(1)) in table_pics and join_pictures(
                                    c.figures[int(prev.group(1))], c.figures.pop(k)):
                                # the table ran on into the next column / page: one picture
                                print(f"[tables] page {pno + 1}: the two parts of a table "
                                      "that broke across columns are joined")
                            else:
                                table_pics.add(k)
                                c.lines.append(token)

                            print(f"[tables] page {pno + 1}: a table with drawings in its "
                                  "cells is kept as a picture")
                        else:
                            c.lines += question_table(body) or [body]
                        c.lines += tail
                    elif p.get("type") == "table_body":  # no structure: keep as picture
                        c.lines.append(figure(pno, p["bbox"]))
                    elif p.get("content"):
                        c.lines.append(_spans_to_html(p["content"]))
            elif t in ("equation", "interline_equation"):
                c.lines.append(f"$${html.escape(clean_latex(b.get('content', '')))}$$")
            elif t in TEXT_LAYER_BLOCKS and doc[pno].get_text("text").strip():
                # answer-key grids get misread as an index: the PDF's own text is exact
                page_ = doc[pno]
                W, H = page_.rect.width, page_.rect.height
                bb = b["bbox"]
                rows = text_layer_rows(page_, fitz.Rect(bb[0] * W, bb[1] * H,
                                                        bb[2] * W, bb[3] * H) + (-2, -2, 2, 2))
                c.lines += [html.escape(r) for r in rows]
            elif t in ("title", "paragraph_title", "doc_title"):
                s = _spans_to_html(b.get("content", ""))
                plain = re.sub(r"<[^>]+>", "", s)
                if s and QNUM_RE.match(plain + " "):
                    c.lines.append(plain)  # a question the layout model took for a heading
                elif s:
                    c.lines.append(f"<h2>{s}</h2>")
            else:  # text, list, caption, ...
                s = _spans_to_html(b.get("content", ""))
                r = rect(bb)
                under = None
                if s and fig_row is not None and len(fig_row["rects"]) > 1 and len(s) <= 40 \
                        and not QNUM_RE.match(s + " ") and not OPT_LABEL_RE.match(_masked(s)):
                    bottom = max(o.y1 for o in fig_row["rects"])
                    under = next((i for i, o in enumerate(fig_row["rects"])
                                  if o.x0 - 6 <= (r.x0 + r.x1) / 2 <= o.x1 + 6
                                  and bottom - 8 <= r.y0 <= bottom + 30), None)
                if under is not None:  # "Figure (b)" under the second figure of a row
                    _row_caption(c.lines, fig_row, under, s)
                elif s:
                    c.lines.append(s)
        if doc[pno].get_text("text").strip():  # put back what MinerU's layout skipped
            for mkey, extra, area in reversed(missed_text(doc[pno], page["blocks"], key)):
                ends = [k[1] for k in keyed[1:]] + [len(c.lines)]
                # MinerU fragments inside the restored run (e.g. only its formula) go
                inside = [(k[1], e) for k, e in zip(keyed, ends)
                          if (area + (-4, -4, 4, 4)).contains(k[2]) and e > k[1]]
                for a, e in reversed(inside):
                    del c.lines[a:e]
                    for k in keyed:
                        if k[1] >= e:
                            k[1] -= e - a
                at = page_start
                for k, e in zip(keyed, [k[1] for k in keyed[1:]] + [len(c.lines)]):
                    if k[0] <= mkey:
                        at = max(at, e)  # after the last block that reads before the run
                c.lines[at:at] = extra
                for k in keyed:
                    if k[1] >= at and k[0] > mkey:
                        k[1] += len(extra)
                print(f"[repair] page {pno + 1}: restored {len(extra)} line(s) MinerU missed: "
                      f"{extra[0][:50]!r}")
    c.lines = [part for ln in c.lines for part in paper_format.FORMAT.split(ln)]
    if any(p.get_text("text").strip() for p in doc):
        c.lines = fix_glued_words(c.lines, doc)

    print(f"[figures] {len(c.figures)} figure(s) cropped at {dpi} dpi; "
          f"{sum(1 for x in moved if x > 3)} box(es) corrected by more than 3 pt"
          + (f"; {sum(1 for f in c.figures.values() if f.svg)} also as vector art"
             if vector else ""))
    return c


# --------------------------------------------------------------------------- #
# 1b. Mathpix (cloud)
# --------------------------------------------------------------------------- #
def mathpix_convert(pdf: Path, cache: Path, app_id: str | None, app_key: str | None,
                    poll: float = 3.0, timeout: float = 1800) -> None:
    md, lines = cache / "paper.md", cache / "lines.json"
    if md.exists() and lines.exists():
        print(f"[mathpix] using cached results in {cache}")
        return
    if not (app_id and app_key):
        sys.exit("Set MATHPIX_APP_ID and MATHPIX_APP_KEY (https://console.mathpix.com).")
    cache.mkdir(exist_ok=True)
    hdr = {"app_id": app_id, "app_key": app_key}
    opts = {"conversion_formats": {"docx": True, "md": True},
            "math_inline_delimiters": ["$", "$"], "math_display_delimiters": ["$$", "$$"],
            "rm_spaces": True}
    with pdf.open("rb") as f:
        r = requests.post(f"{API}/pdf", headers=hdr,
                          files={"file": (pdf.name, f, "application/pdf")},
                          data={"options_json": json.dumps(opts)}, timeout=600)
    r.raise_for_status()
    body = r.json()
    if "pdf_id" not in body:
        sys.exit(f"[mathpix] upload failed: {body}")
    pid = body["pdf_id"]
    (cache / "pdf_id.txt").write_text(pid)
    print(f"[mathpix] uploaded, pdf_id={pid}")
    deadline = time.time() + timeout
    while True:
        s = requests.get(f"{API}/pdf/{pid}", headers=hdr, timeout=60).json()
        status = s.get("status")
        print(f"[mathpix] {status} {s.get('percent_done', '')}%   ", end="\r", flush=True)
        if status == "completed":
            break
        if status == "error" or time.time() > deadline:
            sys.exit(f"\n[mathpix] failed: {s}")
        time.sleep(poll)
    print()
    while time.time() < deadline:  # wait for the md/docx conversions
        cs = requests.get(f"{API}/converter/{pid}", headers=hdr, timeout=60).json() \
            .get("conversion_status", {})
        if all(cs.get(k, {}).get("status") in ("completed", "error") for k in ("md", "docx")):
            break
        time.sleep(poll)

    def fetch(ext: str, dest: Path) -> bool:
        r = requests.get(f"{API}/pdf/{pid}.{ext}", headers=hdr, timeout=600)
        if r.ok and r.content:
            dest.write_bytes(r.content)
            return True
        return False

    fetch("docx", cache / "mathpix_raw.docx")
    if not fetch("lines.json", lines):
        sys.exit("[mathpix] could not download lines.json")
    if not (fetch("md", md) or fetch("mmd", md)):  # md written last = cache complete
        sys.exit("[mathpix] could not download markdown")


def _mathpix_page(url: str, q: dict, pages: dict) -> int | None:
    """Page of a Mathpix crop: match its box against lines.json, else the -N.jpg suffix."""
    for pno, (pw, _, regions) in pages.items():
        tol = 0.01 * pw
        for g in regions:
            if all(abs(float(g.get(k, -1e9)) - q[k]) <= tol
                   for k in ("top_left_x", "top_left_y", "width", "height")):
                return pno
    m = re.search(r"-(\d+)\.(?:jpe?g|png)$", urlparse(url).path)
    return int(m.group(1)) if m and int(m.group(1)) in pages else None


def mathpix_content(cache: Path, doc: fitz.Document, dpi: int, drop_repeated: bool) -> Content:
    data = json.loads((cache / "lines.json").read_text(encoding="utf-8"))
    pages = {int(p["page"]): (float(p["page_width"]), float(p["page_height"]),
                              [ln["region"] for ln in p.get("lines", [])
                               if isinstance(ln.get("region"), dict)])
             for p in data.get("pages", [])}
    md = (cache / "paper.md").read_text(encoding="utf-8").replace("\r\n", "\n")
    md = re.sub(r"\\\[(.+?)\\\]", lambda m: f"$${m.group(1).strip()}$$", md, flags=re.S)
    md = re.sub(r"\\\((.+?)\\\)", lambda m: f"${m.group(1).strip()}$", md, flags=re.S)
    md = re.sub(r"(?<!\$)\$(?!\$)\s*([^$\n]+?)\s*\$(?!\$)", lambda m: f"${m.group(1)}$", md)
    fig_dir = cache / "figures"
    fig_dir.mkdir(exist_ok=True)
    c = Content([], fmt="markdown")

    def repl(m: re.Match) -> str:
        url = m.group(2)
        k = len(c.figures) + 1
        q = {kk: float(v[0]) for kk, v in parse_qs(urlparse(url).query).items()
             if kk in ("height", "width", "top_left_x", "top_left_y")}
        pno = _mathpix_page(url, q, pages) if len(q) == 4 else None
        if pno and pno <= doc.page_count:
            pw, ph, _ = pages[pno]
            sx, sy = doc[pno - 1].rect.width / pw, doc[pno - 1].rect.height / ph
            r = fitz.Rect(q["top_left_x"] * sx, q["top_left_y"] * sy,
                          (q["top_left_x"] + q["width"]) * sx, (q["top_left_y"] + q["height"]) * sy)
            c.figures[k] = crop(doc, pno - 1, r, fig_dir / f"fig{k:03d}.png", dpi)
        else:  # fall back to Mathpix's own (lower-res) crop
            try:
                r = requests.get(url, timeout=60)
                r.raise_for_status()
            except requests.RequestException as e:
                print(f"[figures] could not fetch {url}: {e}")
                return ""
            path = fig_dir / f"fig{k:03d}.jpg"
            path.write_bytes(r.content)
            c.figures[k] = Figure(path, None)
        return IMG.format(k)

    md = MD_IMG_RE.sub(repl, md)
    lines = [LOOSE_LABEL_RE.sub(lambda m: f"({m.group(1) or m.group(2)})", ln)
             for ln in md.split("\n")]
    if drop_repeated:  # running headers/footers repeated on every page
        counts = Counter(ln.strip() for ln in lines if ln.strip())
        junk = {t for t, n in counts.items()
                if n >= 3 and len(t) < 80 and "$" not in t and "|" not in t
                and not t.startswith(("#", "⟪")) and not QNUM_RE.match(t)
                and not OPT_LABEL_RE.match(t)}
        lines = [ln for ln in lines if ln.strip() not in junk]
    c.lines = lines
    print(f"[figures] {len(c.figures)} figure(s) placed")
    return c


# --------------------------------------------------------------------------- #
# MCQ options: split from the stem, repair, lay out in columns
# --------------------------------------------------------------------------- #
def _masked(s: str) -> str:
    return MASK_RE.sub(lambda m: "\0" * len(m.group()), s)


def _family(label: str) -> str:
    return next(f for f in OPT_FAMILIES if label in f)


def _sequence(hits: list[re.Match]) -> list[re.Match]:
    """Longest run of labels continuing (1)(2)(3)... from the first hit."""
    keep = [hits[0]]
    fam = _family(hits[0].group(1))
    for h in hits[1:]:
        if h.group(1) in fam and fam.index(h.group(1)) == fam.index(keep[-1].group(1)) + 1:
            keep.append(h)
    return keep


def split_items(line: str) -> list[str]:
    """'... I. a II. b III. c' / '... A. x B. y C. z' -> one line per item (3+ in order)."""
    hits = list(ITEM_RE.finditer(_masked(line)))
    for fam in ITEM_FAMILIES:
        for i, h in enumerate(hits):
            if h.group(1) != fam[0]:
                continue
            run = [h]
            for g in hits[i + 1:]:
                if len(run) < len(fam) and g.group(1) == fam[len(run)]:
                    run.append(g)
            if len(run) >= 3:
                cuts = [0] + [g.start() for g in run] + [len(line)]
                return [line[a:b].strip() for a, b in zip(cuts, cuts[1:]) if line[a:b].strip()]
    return [line]


def split_stem(line: str) -> list[str]:
    """'12. stem ... (1) a (2) b (3) c' -> ['12. stem ...', '(1) a (2) b (3) c']."""
    hits = list(OPT_LABEL_RE.finditer(_masked(line)))
    for i, h in enumerate(hits):
        if h.start() == 0 or _family(h.group(1)).index(h.group(1)) != 0:
            continue
        if _family(h.group(1)) is OPT_FAMILIES[3]:  # "(i)" mid-sentence is a reference
            continue
        if len(_sequence(hits[i:])) >= 3:
            return [line[:h.start()].rstrip(), line[h.start():]]
    return [line]


def _split_options(line: str) -> list[str] | None:
    """'(1) a (2) b' -> ['(1) a', '(2) b'] when the line starts with an option label."""
    hits = list(OPT_LABEL_RE.finditer(_masked(line)))
    lead = len(line) - len(line.lstrip())
    if not hits or hits[0].start() != lead:
        return None
    bounds = [h.start() for h in _sequence(hits)] + [len(line)]
    return [line[a:b].strip() for a, b in zip(bounds, bounds[1:])]


def _label_hits(line: str) -> list[tuple[int, int]]:
    """[(position, index in its family)] of every option label on the line."""
    return [(h.start(), _family(h.group(1)).index(h.group(1)))
            for h in OPT_LABEL_RE.finditer(_masked(line))
            if _family(h.group(1)) is not OPT_FAMILIES[3]]


def _label_index(opt: str) -> tuple[str, int]:
    lab = OPT_LABEL_RE.match(_masked(opt)).group(1)
    return _family(lab), _family(lab).index(lab)


def _is_object_line(s: str) -> bool:
    """A paragraph that is only a figure or only a display equation."""
    s = s.strip()
    return bool(IMG_TOKEN_RE.fullmatch(s) or re.fullmatch(r"\$\$.+\$\$", s, re.S))


def _inline(s: str) -> str:
    s = s.strip()
    return f"$\\displaystyle {s[2:-2].strip()}$" if s.startswith("$$") else s


def repair_options(lines: list[str]) -> list[str]:
    """Split stems from options; glue '(1)' + figure; label orphan equations in a run."""
    out = []
    for ln in lines:
        if ln.lstrip().startswith("<table"):
            out.append(ln)
            continue
        # Statement-I / Assertion / "Choose the..." on their own lines, then options
        pieces = [s for s in STATEMENT_RE.split(ln) if s.strip()] or [ln]
        if len(pieces) > 1 and re.fullmatch(r"\s*(?:<[^>]+>)*\s*\*?\d{1,3}\.\s*(?:<[^>]+>)*\s*",
                                            pieces[0]):
            pieces[:2] = [pieces[0] + " " + pieces[1]]  # "60. Statement I. ..." keeps its number
        pieces = [x for p in pieces for x in split_items(p)]
        for part in pieces:
            out += split_stem(part)
    # "N. stem (1) a (2) b" + "(3) c" + "(4) d" on the next lines: the run of labels
    # continues across lines, so the stem still gets its options cut off
    for i, ln in enumerate(out):
        if not QNUM_RE.match(re.sub(r"<[^>]+>", "", ln)) or _split_options(ln):
            continue
        hits = [h for h in OPT_LABEL_RE.finditer(_masked(ln)) if h.start() > 0]
        first = next((h for h in hits if _family(h.group(1)).index(h.group(1)) == 0
                      and _family(h.group(1)) is not OPT_FAMILIES[3]), None)
        if first is None:
            continue
        run = _sequence(hits[hits.index(first):])
        nxt = _split_options(out[i + 1]) if i + 1 < len(out) else None
        if nxt and _label_index(nxt[0]) == (_family(first.group(1)), len(run)):
            out[i:i + 1] = [ln[:first.start()].rstrip(), ln[first.start():]]
    # "80. Which ... carbon? (1) (2)" followed by "(1) [fig]", "(2) [fig]": the labels
    # at the end of the stem are OCR's copy of the ones the figures now carry
    for i, ln in enumerate(out):
        tail = re.search(r"(?:\s*\(\s*[1-5]\s*\))+\s*$", ln)
        if not tail or tail.start() == 0:
            continue
        labs = re.findall(r"\(\s*([1-5])\s*\)", tail.group())
        follow = [re.match(r"\s*\(\s*([1-5])\s*\)\s*⟪IMG", x) for x in out[i + 1:i + 1 + len(labs)]]
        if all(follow) and [f.group(1) for f in follow] == labs:
            out[i] = ln[:tail.start()]
    lines, out, i = out, [], 0
    while i < len(lines):
        ln = lines[i]
        nxt = next((j for j in range(i + 1, len(lines)) if lines[j].strip()), None)
        if LABEL_ONLY_RE.match(ln) and nxt is not None and _is_object_line(lines[nxt]):
            out.append(f"{ln.strip()} {_inline(lines[nxt])}")  # '(1)' + its figure/equation
            i = nxt + 1
            continue
        out.append(ln)
        i += 1
    # "51." printed beside its reaction scheme: number and figure share the line
    for i in range(len(out) - 1):
        if re.fullmatch(r"\s*\*?\s*\d{1,3}\.\s*", out[i]) and IMG_TOKEN_RE.fullmatch(out[i + 1].strip()):
            out[i:i + 2] = [f"{out[i].strip()} {out[i + 1].strip()}", ""]
    # options read down the columns of a 2-up grid: "(A) x (C) y" + "(B) z (D) w"
    for i in range(len(out) - 1):
        a, b = _label_hits(out[i]), _label_hits(out[i + 1])
        if len(a) == len(b) == 2 and a[0][0] == 0 and b[0][0] == 0                 and [k for _, k in a] == [0, 2] and [k for _, k in b] == [1, 3]                 and _family(OPT_LABEL_RE.match(_masked(out[i])).group(1))                 is _family(OPT_LABEL_RE.match(_masked(out[i + 1])).group(1)):
            x, y = out[i], out[i + 1]
            out[i:i + 2] = [x[:a[1][0]].strip(), y[:b[1][0]].strip(),
                            x[a[1][0]:].strip(), y[b[1][0]:].strip()]
    fig_opt = re.compile(r"\s*\(([1-5]|[a-eA-E])\)\s*⟪IMG:\d+⟫\s*$")
    i = 0
    while i < len(out):  # "(A) fig", "(C) fig", "(B) fig", "(D) fig" -> A B C D
        j = i
        while j < len(out) and fig_opt.match(out[j]):
            j += 1
        labs = [fig_opt.match(x).group(1) for x in out[i:j]]
        if len(labs) > 1 and sorted(labs) == list(_family(labs[0])[:len(labs)]) != labs:
            out[i:j] = sorted(out[i:j], key=lambda x: fig_opt.match(x).group(1))
        i = max(j, i + 1)
    # an option run "(1) ..", <eq>, <eq>, "(4) .." -> the equations are (2), (3)
    i = 0
    while i < len(out):
        opts = _split_options(out[i]) if out[i].strip() else None
        if opts:
            fam, k = _label_index(opts[-1])
            j = i + 1
            while j < len(out) and (not out[j].strip() or _is_object_line(out[j])):
                j += 1
            orphans = [x for x in range(i + 1, j) if out[x].strip()]
            nxt = _split_options(out[j]) if j < len(out) else None
            if orphans and nxt and _label_index(nxt[0]) == (fam, k + len(orphans) + 1):
                for n, x in enumerate(orphans, 1):
                    out[x] = f"({fam[k + n]}) {_inline(out[x])}"
        i += 1
    return out


FUNC_NAMES = {"sin", "cos", "tan", "cot", "sec", "csc", "ln", "log", "exp", "lim", "max", "min"}


def _tex_width(tex: str) -> float:
    """Printed width of LaTeX in characters (fractions stack, commands collapse)."""
    tex = tex.strip("$")
    tex = re.sub(r"\\(?:mathrm|textrm|text|boldsymbol|mathbf|mathit|displaystyle|left|right)\b",
                 "", tex)
    tex = re.sub(r"\\[,;:! ]|~|\\q?quad", "", tex)
    frac = re.compile(r"\\[dt]?frac\s*\{([^{}]*)\}\s*\{([^{}]*)\}")
    script = re.compile(r"[\^_]\s*\{([^{}]*)\}")
    prev = None
    while prev != tex:  # reduce innermost groups until nothing changes
        prev = tex
        tex = script.sub(lambda m: "x" * int(len(m.group(1).replace(" ", "")) * 0.6 + 0.5), tex)
        tex = frac.sub(lambda m: "x" * int(max(_tex_width(m.group(1)), _tex_width(m.group(2))) + 1),
                       tex)
    tex = re.sub(r"\\([A-Za-z]+)", lambda m: m.group(1) if m.group(1) in FUNC_NAMES else "x", tex)
    return len(re.sub(r"[{}\s^_]", "", tex)) * 1.15


def _visible_len(s: str, figs: dict[int, Figure], char_pt: float) -> float:
    """Rough printed width of a line in characters."""
    extra = 0.0

    def one(m: re.Match) -> str:
        nonlocal extra
        g = m.group()
        if g.startswith("<"):
            return ""
        im = IMG_TOKEN_RE.fullmatch(g)
        if im:
            extra += (figs[int(im.group(1))].width_pt or 120) / char_pt
        else:
            extra += _tex_width(html.unescape(g))
        return ""
    return len(MASK_RE.sub(one, s)) + extra


def layout_options(lines: list[str], figs: dict[int, Figure], lay: Layout,
                   char_pt: float, indent: float, house: bool = False,
                   grid: dict[int, int] | None = None) -> list[str]:
    """Collect consecutive (1)..(4) options and lay them out 4-up, 2-up or 1-up.

    house=True: the template's grid instead - <tab>(1)<tab>value<tab>(2)<tab>value
    rows, 2-up when every value fits its cell, else one option per row.
    """
    col_chars = (lay.col_width - indent) / char_pt
    out, i, qno = [], 0, None
    while i < len(lines):
        qm = QNUM_RE.match(re.sub(r"<[^>]+>", "", lines[i]) + " ")  # "51." alone, too
        if qm:
            qno = int(qm.group(2))
        opts = _split_options(lines[i])
        if not opts or _label_index(opts[0])[1] != 0:
            out.append(lines[i]); i += 1
            continue
        fam, j = _label_index(opts[0])[0], i + 1
        while j < len(lines):  # absorb following option lines (blank lines allowed)
            if not lines[j].strip():
                j += 1; continue
            nxt = _split_options(lines[j])
            if not nxt or _label_index(nxt[0]) != (fam, _label_index(opts[-1])[1] + 1):
                break
            opts += nxt; j += 1
        if len(opts) < 2:
            out.append(lines[i]); i += 1
            continue
        if house:
            values = [re.sub(r"^\s*\([^)]{1,3}\)\s*", "", o).strip() for o in opts]

            def width(v: str) -> float:  # points; a lone figure may shrink to 75 %
                im = IMG_TOKEN_RE.fullmatch(v)
                if im and figs[int(im.group(1))].width_pt:
                    return figs[int(im.group(1))].width_pt * 0.75
                return _visible_len(v, figs, char_pt) * char_pt
            per = 2 if max(map(width, values)) <= house_style.TWO_UP_CELL                 and not len(opts) % 2 else 1
            if grid and qno in grid and _label_index(opts[0])[0] in OPT_FAMILIES[:3]:
                # the source's own arrangement; the template's tab grid holds 2 per row
                per = 1 if len(opts) % 2 else min(grid[qno], 2)
            cell = house_style.TWO_UP_CELL if per == 2 else lay.col_width - house_style.ONE_UP_START
            for o in opts:
                for im in IMG_TOKEN_RE.finditer(o):
                    figs[int(im.group(1))].max_pt = cell - 4
            out += house_style.option_rows(opts, per)
            i = j
            continue
        longest = max(_visible_len(o, figs, char_pt) for o in opts)
        per = 4 if longest * 4 + 8 <= col_chars else 2 if longest * 2 + 4 <= col_chars else 1
        if grid and qno in grid and _label_index(opts[0])[0] in OPT_FAMILIES[:3]:
            per = grid[qno]  # the source's own arrangement
        if per == 4 and len(opts) % 4:
            per = 2
        if per == 2 and len(opts) % 2:
            per = 1
        cell = (lay.col_width - indent) / per - 8
        for o in opts:  # figures inside options must fit their cell
            for im in IMG_TOKEN_RE.finditer(o):
                figs[int(im.group(1))].max_pt = cell - 20
        out.append("")
        for k in range(0, len(opts), per):
            out += [TAB.join(opts[k:k + per]), ""]
        i = j
    return out


# --------------------------------------------------------------------------- #
# 3. Pandoc
# --------------------------------------------------------------------------- #
def pandoc_exe() -> str:
    try:
        import pypandoc
        return pypandoc.get_pandoc_path()
    except Exception:
        pass
    exe = shutil.which("pandoc")
    if not exe:
        sys.exit("Pandoc not found: pip install pypandoc_binary")
    return exe


def render(c: Content, lay: Layout) -> str:
    def img(m: re.Match) -> str:
        f = c.figures[int(m.group(1))]
        w = f.width_in or None
        w = min(w, lay.col_width / 72) if w else None
        if c.fmt == "markdown":
            return f"![]({f.path.as_posix()})" + (f"{{width={w:.2f}in}}" if w else "")
        return f'<img src="{html.escape(f.path.as_posix())}"' + \
            (f' style="width:{w:.2f}in"' if w else "") + ">"

    if c.fmt == "markdown":
        lines = [QNUM_RE.sub(r"\1\2\\.", ln) for ln in c.lines]  # no Word auto-lists
        return IMG_TOKEN_RE.sub(img, "\n".join(lines))
    body = []
    for ln in c.lines:
        if not ln.strip():
            continue
        ln = IMG_TOKEN_RE.sub(img, ln)
        body.append(ln if ln.lstrip().startswith(("<table", "<h2>")) else f"<p>{ln}</p>")
    return "<!doctype html><html><head><meta charset='utf-8'></head><body>\n" \
        + "\n".join(body) + "\n</body></html>"


def make_reference_docx(pandoc: str, dest: Path, font: str, size: float) -> None:
    subprocess.run([pandoc, "-o", str(dest), "--print-default-data-file", "reference.docx"],
                   check=True)
    d = Document(dest)
    for st in d.styles:
        if st.type != 1:  # paragraph styles only
            continue
        st.font.name = font
        rpr = st.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts"); rpr.append(fonts)
        for a in ("w:ascii", "w:hAnsi", "w:cs"):
            fonts.set(qn(a), font)
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme"):
            fonts.attrib.pop(qn(a), None)
        if st.name.startswith(("Heading", "Title")):
            st.font.size = Pt(size + 1)
            st.font.bold = True
            st.font.color.rgb = None
            st.font.underline = True
            st.paragraph_format.alignment = 1  # centred, like "ANSWER KEY"
        elif st.name in ("Normal", "Body Text", "First Paragraph", "Compact", "Figure",
                         "Captioned Figure", "Image Caption", "Table"):
            st.font.size = Pt(size)
        pf = st.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(3), 1.0
    d.save(dest)


def run_pandoc(pandoc: str, src: Path, fmt: str, ref: Path, out: Path) -> None:
    if fmt == "markdown":
        fmt = ("markdown+tex_math_dollars+tex_math_single_backslash+pipe_tables+grid_tables"
               "+link_attributes-fancy_lists-example_lists-citations-smart-subscript-superscript")
    else:
        fmt = "html+tex_math_dollars"
    subprocess.run([pandoc, str(src), "-f", fmt, "-t", "docx", "--reference-doc", str(ref),
                    "--resource-path", str(src.parent), "-o", str(out)], check=True)


# --------------------------------------------------------------------------- #
# 4. Word layout pass
# --------------------------------------------------------------------------- #
def _set_child(parent, tag: str, before: tuple[str, ...] = ()):
    """Get or create parent/<tag>, inserted before the first of `before` (schema order)."""
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        anchor = next((parent.find(qn(b)) for b in before if parent.find(qn(b)) is not None), None)
        anchor.addprevious(el) if anchor is not None else parent.append(el)
    return el


def fit_table(tbl, width_pt: float, indent_pt: float) -> None:
    """Shrink a table to the column width (columns keep their proportions) + grid borders."""
    later = ("w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription")
    pr = tbl.tblPr
    w = _set_child(pr, "w:tblW", ("w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders",
                                   "w:shd") + later)
    w.set(qn("w:w"), str(int(width_pt * 20)))
    w.set(qn("w:type"), "dxa")
    ind = _set_child(pr, "w:tblInd", ("w:tblBorders", "w:shd") + later)
    ind.set(qn("w:w"), str(int(indent_pt * 20)))
    ind.set(qn("w:type"), "dxa")
    borders = _set_child(pr, "w:tblBorders", ("w:shd",) + later)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = _set_child(borders, f"w:{side}")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0"); b.set(qn("w:color"), "000000")
    _set_child(pr, "w:tblLayout", later[1:]).set(qn("w:type"), "fixed")
    cols = tbl.tblGrid.findall(qn("w:gridCol"))
    need = [6] * len(cols)  # widest single-span cell text per grid column, in chars
    for tr in tbl.findall(qn("w:tr")):
        k = 0
        for tc in tr.findall(qn("w:tc")):
            span = tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
            n = int(span.get(qn("w:val"))) if span is not None else 1
            if n == 1 and k < len(need):
                text = sum(len(t.text or "") for t in tc.iter(qn("w:t")))
                math = sum(len(t.text or "") for t in tc.iter(qn("m:t")))
                math += 2 * len(list(tc.iter(qn("m:rad"), qn("m:f"), qn("m:sSup"))))
                need[k] = max(need[k], min(30, text + 1.4 * math + 4))  # + cell padding
            k += n
    total = width_pt * 20
    new = [int(total * x / sum(need)) for x in need]
    # short columns ("(A)", "(III)", "0.25") must not wrap: give them their full
    # width (11 pt text ~6 pt/char + cell margins), taken from the widest column
    for k, x in enumerate(need):
        floor = int(((x - 4) * 6 + 14) * 20)
        if x <= 9 and new[k] < floor:
            widest = max(range(len(new)), key=lambda i: new[i])
            give = min(floor - new[k], new[widest] - floor)
            if widest != k and give > 0:
                new[k] += give
                new[widest] -= give
    for c, n in zip(cols, new):
        c.set(qn("w:w"), str(n))
    for tr in tbl.findall(qn("w:tr")):  # cell widths follow their grid span
        k = 0
        for tc in tr.findall(qn("w:tc")):
            tcpr = tc.get_or_add_tcPr()
            span = tcpr.find(qn("w:gridSpan"))
            n = int(span.get(qn("w:val"))) if span is not None else 1
            tcw = _set_child(tcpr, "w:tcW", ("w:gridSpan", "w:vMerge", "w:tcBorders", "w:shd",
                                             "w:noWrap", "w:tcMar", "w:vAlign"))
            tcw.set(qn("w:w"), str(sum(new[k:k + n])))
            tcw.set(qn("w:type"), "dxa")
            k += n


def fit_figure_rows(lines: list[str], figs: dict[int, Figure], lay: Layout,
                    indent: float) -> None:
    """Figures printed side by side share the column: one tab cell each."""
    for ln in lines:
        cells = ln.split(TAB)
        if len(cells) > 1 and all(IMG_TOKEN_RE.fullmatch(x.strip()) for x in cells):
            for x in cells:
                f = figs[int(IMG_TOKEN_RE.fullmatch(x.strip()).group(1))]
                f.max_pt = min(f.max_pt or 1e9, (lay.col_width - indent) / len(cells) - 6)


SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def attach_svgs(docx_path: Path, figs: dict[int, Figure]) -> int:
    """Give every line drawing its vector twin: Word 2019 / 365 shows the SVG, older
    programs the PNG beside it. Returns the number of pictures upgraded."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part
    from lxml import etree

    by_name = {f.path.name: f for f in figs.values() if f.svg and f.svg.exists()}
    if not by_name:
        return 0
    d = Document(docx_path)
    rids: dict[str, str] = {}
    n = 0
    for pic in d.element.body.iter(qn("pic:pic")):
        cnv = next(pic.iter(qn("pic:cNvPr")), None)
        blip = next(pic.iter(f"{{{A_NS}}}blip"), None)
        f = by_name.get(Path(cnv.get("descr", "")).name) if cnv is not None else None
        if f is None or blip is None or blip.find(f"{{{A_NS}}}extLst") is not None:
            continue
        if f.path.name not in rids:
            part = Part(PackURI(f"/word/media/{f.svg.name}"), "image/svg+xml",
                        f.svg.read_bytes(), d.part.package)
            rids[f.path.name] = d.part.relate_to(part, RT.IMAGE)
        ext_lst = etree.SubElement(blip, f"{{{A_NS}}}extLst")
        ext = etree.SubElement(ext_lst, f"{{{A_NS}}}ext", uri=SVG_EXT)
        svg = etree.SubElement(ext, f"{{{SVG_NS}}}svgBlip", nsmap={"asvg": SVG_NS})
        svg.set(f"{{{R_NS}}}embed", rids[f.path.name])
        n += 1
    d.save(docx_path)
    return n


def keep_questions_together(d) -> None:

    """Plain layout: every "N.<tab>stem" block up to the next number stays in one
    column (the answer key and solutions that follow are left to flow)."""
    blocks, block = [], None
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            style = el.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
            if style is not None and (style.get(qn("w:val")) or "").startswith(("Heading", "Title")):
                block = None  # "CHEMISTRY" starts afresh
                if re.search(r"ANSWER|SOLUTION|HINT", house_style._text(el), re.I):
                    break
                continue
            if re.match(r"\s*\*?\s*\d{1,3}\.\t", house_style._full_text(el)):
                block = []
                blocks.append(block)
        elif el.tag != qn("w:tbl"):
            continue
        if block is not None:
            block.append(el)
    for block in blocks:
        house_style.keep_together(block)


def apply_layout(docx_path: Path, lay: Layout, indent_pt: float,
                 figs: dict[int, Figure] | None = None) -> None:
    d = Document(docx_path)
    by_name = {f.path.name: f for f in (figs or {}).values()}
    for sec in d.sections:
        sec.page_width, sec.page_height = Pt(lay.width), Pt(lay.height)
        sec.left_margin, sec.right_margin = Pt(lay.left), Pt(lay.right)
        sec.top_margin, sec.bottom_margin = Pt(lay.top), Pt(lay.bottom)
        sp = sec._sectPr
        cols = sp.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            sp.append(cols)
        cols.set(qn("w:num"), str(lay.columns))
        cols.set(qn("w:space"), str(int(lay.gutter * 20)))
        if lay.columns > 1:
            cols.set(qn("w:sep"), "1")
        for tag in ("w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
                    "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings"):
            nxt = sp.find(qn(tag))  # schema order: w:cols precedes these
            if nxt is not None:
                nxt.addprevious(cols)
                break

    for p in d.paragraphs:  # ANSWER KEY / SOLUTIONS start a new page, as in the source
        if p.style.name.startswith("Heading") and p._p.getprevious() is not None \
                and re.search(r"ANSWER|SOLUTION|HINT", p.text, re.I):  # not "CHEMISTRY"
            p.paragraph_format.page_break_before = True
    ind, col = Pt(indent_pt), Pt(lay.col_width)
    in_paper = False  # past the first numbered line
    for tbl in d.tables:
        fit_table(tbl._tbl, lay.col_width - indent_pt, indent_pt)
    for p in d.paragraphs:
        runs = p.runs
        n_tabs = sum(r.text.count(TAB) for r in runs)
        if n_tabs:  # option row: real tabs + evenly spaced stops across the column
            for r in runs:
                if TAB in r.text:
                    r.text = r.text.replace(TAB, "\t")
                if r.text.lstrip("\t").startswith("(") or "\t(" in r.text:
                    r.text = paper_format.FORMAT.label_text(r.text)
            p.paragraph_format.left_indent = ind
            step = (col - ind) / (n_tabs + 1)
            for k in range(1, n_tabs + 1):
                p.paragraph_format.tab_stops.add_tab_stop(int(ind + step * k))
            house_style.lower_option_figures(p._p)
            continue
        text = p.text
        pics = [by_name.get(Path(c.get("descr", "")).name) for c in p._p.iter(qn("pic:cNvPr"))]
        if pics and not text.strip():
            # a figure on its own: where the PDF prints it - mid-column, or as far in
            f, width = pics[0], sum(int(e.get("cx")) for e in p._p.iter(qn("wp:extent"))) / 12700
            if f is not None and f.centred and len(pics) == 1:
                p.paragraph_format.alignment = 1
            elif f is not None and f.left_pt is not None:
                p.paragraph_format.left_indent = Pt(max(0, min(f.left_pt, lay.col_width - width)))
            else:
                p.paragraph_format.left_indent = ind
        elif OPT_LABEL_RE.match(text):
            p.paragraph_format.left_indent = ind
            house_style.lower_option_figures(p._p)
            if runs and not paper_format.FORMAT.plain:
                runs[0].text = paper_format.FORMAT.label_text(runs[0].text)
        elif QNUM_RE.match(text) and runs:
            in_paper = True
            # hanging bold number "12.<tab>stem"; wrapped lines align with the stem
            p.paragraph_format.left_indent = ind
            p.paragraph_format.first_line_indent = -ind
            first = runs[0]
            m = re.match(r"^\s*(\*?\s*\d{1,3}\.)\s*", first.text)
            if not m:
                continue
            rest = first.text[m.end():]
            if not rest and len(runs) > 1:
                runs[1].text = runs[1].text.lstrip()
            num = copy.deepcopy(first._r)
            first._r.addprevious(num)
            num_run = type(first)(num, p)
            digits = re.search(r"\d+", m.group(1)).group()
            num_run.text = paper_format.FORMAT.number_text(
                digits, "*" if "*" in m.group(1) else "") + "\t"
            num_run.bold = True
            if rest:
                first.text = rest
            else:
                p._p.remove(first._r)
            if pics and not text[m.end():].strip():  # "51." beside its reaction scheme
                house_style.lower_option_figures(p._p)
        elif in_paper and (pics or text.strip()) and not p.style.name.startswith(
                ("Heading", "Title")):
            p.paragraph_format.left_indent = ind  # "Give the structure of X" under its stem
    # no picture may be wider than the room its paragraph has in the column
    for p in d.paragraphs:
        room = lay.col_width - (p.paragraph_format.left_indent or 0) / 12700
        for inline in p._p.iter(qn("wp:inline")):
            ext = inline.find(qn("wp:extent"))
            cx, cy = int(ext.get("cx")), int(ext.get("cy"))
            if cx / 12700 <= room + 0.5 or room < 40:
                continue
            k = room * 12700 / cx
            for e in [ext] + list(inline.iter(qn("a:ext"))):
                if e.get("cx") and e.get("cy"):
                    e.set("cx", str(int(int(e.get("cx")) * k)))
                    e.set("cy", str(int(int(e.get("cy")) * k)))
    keep_questions_together(d)
    d.save(docx_path)



GRID_LABEL_RE = re.compile(r"\(([1-4]|[A-D]|[a-d])\)")


# a label after one of these is a reference inside a sentence ("Both (A) and (R)")
LINK_WORDS = {"and", "or", "&", "both", "either", "neither", "nor", "than", "to", "of", "is",
              "are", "in", "only", "but", "with", "for", "from", "by", "as", "statement",
              "statements", "assertion", "reason", "option", "options", "not"}


def pdf_option_grid(doc: fitz.Document) -> dict[int, int]:
    """{question: 4, 2 or 1} - how many options the source PDF prints per row.

    Read from the PDF's text layer in column order, so the output copies the
    editor's own choice of 4-up / 2-up / one-per-row options question by question.
    Numbering may start anywhere (section papers begin at 46, 91, 136 ...).
    """
    grid: dict[int, int] = {}
    cur = 0
    cands: list[tuple] = []  # (family, option index, y) of the label-like words, in order

    def close():
        # the options are the LAST run (1)(2).. / (A)(B).. / (a)(b).. of the question:
        # walk back from its highest label ("Assertion (A)" in the stem comes earlier)
        best = None
        for fam in OPT_FAMILIES[:3]:
            mine = [(i, k, y) for i, (f, k, y) in enumerate(cands) if f is fam]
            for top in sorted({k for _, k, _ in mine if k}, reverse=True):
                seen, at = {}, len(cands)
                for k in range(top, -1, -1):
                    hit = next(((i, y) for i, kk, y in reversed(mine) if kk == k and i < at), None)
                    if hit is None:
                        break
                    at, seen[k] = hit
                if len(seen) == top + 1:
                    end = max(i for i, k, _ in mine if k == top)
                    if best is None or end > best[0]:
                        best = (end, seen)
                    break
        if cur and best:
            seen = best[1]
            same = lambda a, b: abs(seen[a] - seen[b]) < 4  # noqa: E731
            four = 2 in seen and 3 in seen and same(0, 3)
            grid[cur] = 4 if four and same(0, 1) else 2 if same(0, 1) else 1

    for page in doc:
        for col in text_segments(page):
            # 5th percentile, not min: a wide word of the other column can straddle the middle
            left = sorted(w[0] for w in col)[len(col) // 20]
            prev = None
            for w in sorted(col, key=lambda w: (round(w[1]), w[0])):
                t = w[4]
                # an option label opens its row or stands after a tab gap; "(A)" inside
                # a sentence ("Both (A) and (B)") follows a plain space
                opens = prev is None or round(prev[1]) != round(w[1]) or w[0] - prev[2] > 6
                linked = not opens and (prev[4].endswith(",") or prev[4].lower() in LINK_WORDS)
                prev = w
                if t == "ANSWER":  # the key's heading (not "...correct answer...")
                    close()
                    return grid
                m = re.fullmatch(r"\*?(\d{1,3})\.", t)
                # the next number, or (before any question with options) any number:
                # cover instructions "1. 2. 3." must not lock the count
                if m and w[0] - left < 30 and (int(m.group(1)) == cur + 1
                                               or (not grid and not cands)):
                    close()
                    cur, cands = int(m.group(1)), []
                elif cur and not linked and (lab := GRID_LABEL_RE.fullmatch(t)):
                    f = _family(lab.group(1))
                    cands.append((f, f.index(lab.group(1)), w[1]))
    close()
    return grid


_ROWS_CACHE: dict[int, dict[int, list[str]]] = {}


def pdf_question_rows(doc: fitz.Document) -> dict[int, list[str]]:
    """{question: its printed rows} from the text layer, in column/page reading order."""
    if id(doc) not in _ROWS_CACHE:
        _ROWS_CACHE[id(doc)] = _question_rows(doc)
    return _ROWS_CACHE[id(doc)]


BULLETS = "•●▪■◦○➢➤►"


def recover_bullets(lines: list[str], doc: fitz.Document) -> list[str]:
    """Put back the bullet OCR dropped in front of a "•  item" line (text layer decides)."""
    if not any(p.get_text("text").strip() for p in doc):
        return lines
    rows = pdf_question_rows(doc)
    starts = [i for i, ln in enumerate(lines) if QNUM_RE.match(re.sub(r"<[^>]+>", "", ln))]
    fixed = 0
    for a, b in zip(starts, starts[1:] + [len(lines)]):
        n = int(QNUM_RE.match(re.sub(r"<[^>]+>", "", lines[a])).group(2))
        heads = {_norm(r.lstrip(BULLETS))[:12] for r in rows.get(n, [])
                 if r[:1] in BULLETS and len(_norm(r)) >= 12}
        if not heads:
            continue
        for i in range(a + 1, b):
            s = re.sub(r"<[^>]+>", "", lines[i]).strip()
            if s[:1] not in BULLETS and _norm(lines[i])[:12] in heads:
                lines[i] = "• " + lines[i].lstrip()
                fixed += 1
    if fixed:
        print(f"[repair] {fixed} bullet(s) restored from the PDF text layer")
    return lines


def _question_rows(doc: fitz.Document) -> dict[int, list[str]]:
    out: dict[int, list[str]] = {}
    cur, has_options = 0, False
    for page in doc:
        for col in text_segments(page):
            left = sorted(w[0] for w in col)[len(col) // 20]
            rows: list[list] = []
            for w in sorted(col, key=lambda w: (w[1], w[0])):
                if rows and abs(rows[-1][0][1] - w[1]) < 3:
                    rows[-1].append(w)
                else:
                    rows.append([w])
            for r in rows:
                r.sort(key=lambda w: w[0])
                if any(w[4] == "ANSWER" for w in r):
                    return out
                m = re.fullmatch(r"\*?(\d{1,3})\.", r[0][4])
                # the next number - or any number until a question with options was seen
                # (cover instructions "1. 2. 3. 4." come before the paper's own 1.)
                if m and r[0][0] - left < 30 and (int(m.group(1)) == cur + 1 or not has_options):
                    cur = int(m.group(1))
                    if not has_options:
                        out[cur] = []  # drop what an instruction with this number collected
                if cur:
                    row = " ".join(w[4] for w in r).translate(_SYMBOL_MAP)
                    has_options = has_options or row.startswith(("(1)", "(A)", "(a)"))
                    out.setdefault(cur, []).append(row)
    return out


ROW_LABEL_RE = re.compile(r"\(\s*([1-5]|[A-E]|[a-e])\s*\)")


def recover_options(lines: list[str], doc: fitz.Document) -> list[str]:
    """A question whose (1)..(4) options OCR lost or scrambled gets them from the PDF.

    Only questions without any usable option line are touched; True/False grids
    ("(1) F F T F") keep their letters one per tab cell, like the source.
    """
    if not any(p.get_text("text").strip() for p in doc):
        return lines
    rows = None
    starts = [i for i, ln in enumerate(lines) if QNUM_RE.match(re.sub(r"<[^>]+>", "", ln))]
    for a, b in reversed(list(zip(starts, starts[1:] + [len(lines)]))):
        block = lines[a:b]
        if any((o := _split_options(ln)) and _label_index(o[0])[1] == 0
               and all(OPT_LABEL_RE.sub("", x, count=1).strip() for x in o) for ln in block):
            continue  # it has options, with their values
        if rows is None:
            rows = pdf_question_rows(doc)
        n = int(QNUM_RE.match(re.sub(r"<[^>]+>", "", lines[a])).group(2))
        opts, fam, broken = [], None, False
        for r in rows.get(n, []):
            hits = list(ROW_LABEL_RE.finditer(r))
            if not hits or hits[0].start() != 0:
                continue
            fam = fam or _family(hits[0].group(1))
            run = []  # labels continuing (1)(2).. / (A)(B)..; "(A)" inside a value is text
            for h in hits:
                k = len(opts) + len(run)
                if k < len(fam) and h.group(1) == fam[k]:
                    run.append(h)
            if not run or run[0].start() != 0:
                broken = True
                break
            hits = run
            cuts = [h.start() for h in hits] + [len(r)]
            for h, (x, y) in zip(hits, zip(cuts, cuts[1:])):
                value = r[h.end():y].strip()
                if re.fullmatch(r"[TF](?: [TF])+", value):  # True/False grid row
                    value = TAB.join(value.split())
                opts.append(f"({h.group(1)}) {html.escape(value)}")
        if len(opts) < 2 or broken:
            continue
        # option values that are formulas/pictures have no text layer: nothing to gain
        if sum(1 for o in opts if o[3:].strip()) < len(opts) * 0.75:
            continue
        # the scrambled remains: lines of labels ("(1) (2) (3) (4)") or grid letters ("FFTF FFFT")
        keep = [ln for ln in block[1:] if len(list(OPT_LABEL_RE.finditer(_masked(ln)))) < 3
                and not re.fullmatch(r"\s*[TF]{2,}(?:\s+[TF]{2,})*\s*", ln)]
        # OCR's copy of the options as bare formulas ("I -> P; II -> Q", one per option)
        bare = [ln for ln in keep if re.fullmatch(r"\s*\$\$.+\$\$\s*", ln, re.S)]
        if len(bare) == len(opts):
            keep = [ln for ln in keep if ln not in bare]
        lines[a:b] = [block[0]] + keep + opts
        print(f"[repair] question {n}: options rebuilt from the PDF text layer")
    return lines


def _norm(s: str) -> str:
    s = MASK_RE.sub(" ", html.unescape(s))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def recover_headings(lines: list[str], doc: fitz.Document, head_re: re.Pattern,
                     with_answer: bool) -> list[str]:
    """Re-insert 'N.' / 'N. (k)' headings OCR dropped, using the PDF's own text layer.

    For a missing number N, the text layer gives the words that follow "N." (and its
    "(k)"); the converted line starting with those words gets the heading put back.
    """
    def number(ln):
        m = head_re.match(re.sub(r"<[^>]+>", "", ln).strip())
        return int(m.group(1)) if m else None

    nums = [n for n in (number(ln) for ln in lines) if n]
    if not nums:
        return lines
    missing = [n for n in range(1, max(nums)) if n not in nums]
    if not missing:
        return lines
    text = [paper_format.FORMAT.line(ln.strip()).strip()
            for p in doc for ln in p.get_text().split("\n")]
    for n in missing:
        snippet, answer = None, None
        for j, t in enumerate(text):
            if not re.fullmatch(rf"{n}\.", t):
                continue
            k = j + 1
            if with_answer:
                a = re.fullmatch(r"\(\s*([1-5a-dA-D])\s*\)", text[k] if k < len(text) else "")
                if not a:
                    continue
                answer, k = a.group(1), k + 1
            follow = " ".join(x for x in text[k:k + 4] if x)
            if len(_norm(follow)) >= 12:
                snippet = _norm(follow)[:14]
        if not snippet:
            continue
        prev_i = max((i for i, ln in enumerate(lines) if number(ln) and number(ln) < n), default=-1)
        next_i = min((i for i, ln in enumerate(lines) if number(ln) and number(ln) > n),
                     default=len(lines))
        for i in range(prev_i + 1, next_i):
            if _norm(lines[i]).startswith(snippet[:12]):
                head = f"{n}. ({answer})" if with_answer else f"{n}."
                lines = lines[:i] + ([head] if with_answer else []) + \
                    ([f"{head} {lines[i]}"] if not with_answer else [lines[i]]) + lines[i + 1:]
                print(f"[repair] restored heading {head} from the PDF text layer")
                break
    return lines


def template_layout(sk, fallback: Layout) -> Layout:
    """Page geometry of the template's 2-column questions section."""
    sp = None
    if sk.q_close is not None:
        sp = sk.q_close.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
    if sp is None:
        return fallback
    tw = lambda el, a, d=0: int(el.get(qn(a)) or d) / 20 if el is not None else d / 20  # noqa: E731
    sz, mar, cols = sp.find(qn("w:pgSz")), sp.find(qn("w:pgMar")), sp.find(qn("w:cols"))
    return Layout(tw(sz, "w:w", 11906), tw(sz, "w:h", 16838), tw(mar, "w:left", 720),
                  tw(mar, "w:right", 720), tw(mar, "w:top", 720), tw(mar, "w:bottom", 720),
                  int(cols.get(qn("w:num")) or 1) if cols is not None else 1,
                  tw(cols, "w:space", 720))


# --------------------------------------------------------------------------- #
def main() -> None:
    # progress text holds →, ✓, maths symbols; a pipe to the Paper Maker window would
    # otherwise use the ANSI code page (cp1252) and crash on the first such print
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("pdf", type=Path)
    ap.add_argument("-o", "--output", type=Path)
    ap.add_argument("--engine", choices=["mineru", "mathpix"], default="mineru")
    ap.add_argument("--pages", default="all", help="MinerU page range, e.g. 1-3,5 (default all)")
    ap.add_argument("--tier", default="basic", choices=["flash", "basic", "standard"],
                    help="MinerU model tier (default basic; standard adds a slower local VLM)")
    ap.add_argument("--mineru", help="path to mineru-kit.exe")
    ap.add_argument("--columns", choices=["auto", "1", "2"], default="auto")
    ap.add_argument("--dpi", type=int, default=300, help="figure resolution (default 300)")
    ap.add_argument("--vector-figures", action="store_true",
                    help="line drawings also go in as vector art (SVG, Word 2019 / 365): sharp "
                         "at any zoom, and 'Convert to Shape' in Word makes them editable")
    ap.add_argument("--font", default="Times New Roman")
    ap.add_argument("--size", type=float, default=11, help="body font size in pt")
    ap.add_argument("--indent", type=float, default=22, help="question/option indent in pt")
    ap.add_argument("--keep-repeated", action="store_true",
                    help="Mathpix: keep lines repeated on many pages (default: dropped)")
    ap.add_argument("--template", type=Path,
                    help="original Word file of the same series: output copies its cover, "
                         "header/footer, sections and paragraph formatting (default: a .docx "
                         "with the PDF's name next to it, if there is one)")
    ap.add_argument("--no-template", action="store_true", help="ignore any template")
    ap.add_argument("--cache-dir", type=Path,
                    help="where engine results are kept (default: next to the PDF)")
    ap.add_argument("--app-id", default=os.environ.get("MATHPIX_APP_ID"))
    ap.add_argument("--app-key", default=os.environ.get("MATHPIX_APP_KEY"))
    a = ap.parse_args()

    pdf = a.pdf.resolve()
    # never overwrite the original Word file sitting next to the PDF
    out = (a.output or pdf.with_name(pdf.stem + "_editable.docx")).resolve()
    for n in range(2, 30):  # the previous result is still open in Word: write beside it
        try:
            if out.exists():
                with open(out, "r+b"):
                    pass
            break
        except PermissionError:
            base = (a.output or pdf.with_name(pdf.stem + "_editable.docx")).resolve()
            out = base.with_name(f"{base.stem} ({n}){base.suffix}")
            print(f"[output] the file is open in another program; writing {out.name} instead")
    template = None if a.no_template else (a.template or pdf.with_suffix(".docx"))
    if template is not None and not template.exists():
        if a.template:
            sys.exit(f"template not found: {template}")
        template = None
    if template is not None and template.resolve() == out:
        sys.exit("the output would overwrite the template - choose another -o")
    doc = fitz.open(pdf)
    cache_root, key = pdf.parent, ""
    if a.cache_dir:  # shared cache: tell same-named PDFs from different folders apart
        cache_root = a.cache_dir.resolve()
        cache_root.mkdir(parents=True, exist_ok=True)
        st = pdf.stat()
        key = "_" + hashlib.sha1(f"{pdf}|{st.st_size}|{st.st_mtime_ns}".encode()).hexdigest()[:8]

    if a.engine == "mineru":
        pages = parse_pages(a.pages, doc.page_count)
        tag = "" if a.pages == "all" else "_p" + re.sub(r"[^0-9]+", "-", a.pages)
        cache = cache_root / f"{pdf.stem}{key}_mineru{tag}_{a.tier}"
        mineru_convert(pdf, cache, find_mineru(a.mineru), pages, a.tier)
        # the paper's own numbering / option style, from its text (OCR text for a scan)
        own = [ln for p in doc for ln in p.get_text().split("\n")]
        if not any(x.strip() for x in own):
            own = [_spans_html(b.get("content", "")) for pg in json.loads(
                (cache / "middle_json.json").read_text(encoding="utf-8"))["pages"]
                for b in pg["blocks"] if b["type"] not in ("image", "table")]
        paper_format.FORMAT = paper_format.detect(own)
        if not paper_format.FORMAT.plain:
            print(f"[format] this paper writes {paper_format.FORMAT.describe()}; "
                  "read as 12. / (A) and written back in its own style")
        content = mineru_content(cache, doc, a.dpi, pages, a.vector_figures)
    else:
        cache = cache_root / f"{pdf.stem}{key}_mathpix"
        mathpix_convert(pdf, cache, a.app_id, a.app_key)
        content = mathpix_content(cache, doc, a.dpi, drop_repeated=not a.keep_repeated)

    lay = analyse_layout(doc, a.columns)
    print(f"[layout] {lay.width:.0f}x{lay.height:.0f}pt, {lay.columns} column(s), "
          f"gutter {lay.gutter:.0f}pt, column {lay.col_width:.0f}pt")

    char_pt = a.size * 0.5  # ~0.5em average glyph width
    content.lines = repair_options(repair_tables(content.lines))
    sk = None
    if template is not None:
        sk = house_style.load_skeleton(template)
        lay = template_layout(sk, lay)
        print(f"[template] {template.name}: cover, header/footer, sections and paragraph "
              f"formatting copied; column {lay.col_width:.0f}pt")
        parts = house_style.split_parts(content.lines)
        parts["q"] = recover_headings(parts["q"], doc, house_style.QNUM_RE, False)
        parts["sol"] = recover_headings(parts["sol"], doc, house_style.SOL_RE, True)
        parts["q"] = recover_bullets(recover_options(parts["q"], doc), doc)
        grid = pdf_option_grid(doc)
        if grid:
            print(f"[options] 2-up / one-per-row taken from the PDF for {len(grid)} questions")
        parts["q"] = layout_options(parts["q"], content.figures, lay, char_pt, a.indent,
                                    house=True, grid=grid)
        content.lines = house_style.tag_roles(
            parts, lambda ln: ln.startswith(house_style.TAB + "("))
        for f in content.figures.values():  # pictures sit after the 0.39" tab
            f.max_pt = min(f.max_pt or 1e9, lay.col_width - 28.35)
        print(f"[parts] {sum(1 for ln in parts['q'] if QNUM_RE.match(ln))} questions, "
              f"{len(parts['key'])} answer-key lines, "
              f"{sum(1 for ln in parts['sol'] if house_style.SOL_RE.match(ln))} solutions")
    else:
        heads = [(i, int(m.group(2))) for i, ln in enumerate(content.lines)
                 if (m := QNUM_RE.match(re.sub(r"<[^>]+>", "", ln)))]
        begin = heads[house_style.question_start([n for _, n in heads])][0] if heads else 0
        body = content.lines[begin:]
        ends = [i for i, ln in enumerate(body) if ln.startswith("<h2>") and re.search(
            r"ANSWER|SOLUTION|HINT", ln, re.I)]  # numbers repeat after the key's heading
        cut = ends[0] if ends else len(body)
        body = recover_headings(body[:cut], doc, house_style.QNUM_RE, False) + body[cut:]
        content.lines = content.lines[:begin] + recover_bullets(recover_options(body, doc), doc)
        content.lines = layout_options(content.lines, content.figures, lay, char_pt, a.indent,
                                       grid=pdf_option_grid(doc))
    fit_figure_rows(content.lines, content.figures, lay, a.indent)
    src = cache / ("paper.clean.md" if content.fmt == "markdown" else "paper.html")
    src.write_text(render(content, lay), encoding="utf-8")

    pandoc = pandoc_exe()
    ref = cache / "reference.docx"
    make_reference_docx(pandoc, ref, a.font, a.size)
    if sk is not None:
        tmp = cache / "pandoc_body.docx"
        run_pandoc(pandoc, src, content.fmt, ref, tmp)
        house_style.build(sk, tmp, out, fit_table, lay.col_width,
                          frozenset(f.path.name for f in content.figures.values() if f.centred))
    else:
        run_pandoc(pandoc, src, content.fmt, ref, out)
        # a longer number ("Q.12") needs a wider hanging indent
        wide = max(0, len(paper_format.FORMAT.number_text("188")) - 4) * 5.5
        apply_layout(out, lay, a.indent + wide, content.figures)


    if a.vector_figures:
        n_svg = attach_svgs(out, content.figures)
        print(f"[figures] {n_svg} line drawing(s) placed as vector art (PNG kept as fallback)")
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    n_eq = len(re.findall(r"<m:oMath[ >]", xml))
    raw = len(re.findall(r"<w:t[^>]*>[^<]*\\(?:frac|mathrm|sqrt|left)", xml))
    print(f"[done] {out}  ({n_eq} native Word equations"
          + (f", {raw} formula(s) Pandoc could not parse - left as LaTeX text" if raw else "") + ")")


if __name__ == "__main__":
    main()
