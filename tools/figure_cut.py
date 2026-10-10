"""Exact figure crops.

The layout model only *guesses* a figure's box: it can stop short of the real
edge, cut a stroke, or include the last line of the question above. This module
refines the guess on the page itself:

* everything that is known NOT to be the figure is taken away first - the words
  of the text blocks around it (PDF text layer; block areas on a scanned page)
  and the area of the other figures on the page;
* what is left near the guess is grouped into connected pieces of ink (pieces
  closer than ~2.5 pt belong together); a piece that lies mostly inside the
  guess is part of the figure and is taken WHOLE, so nothing is cut, and a piece
  that lies mostly outside is a neighbour and is left out;
* the crop is the tight box around the figure's own pieces; neighbours that still
  fall inside that box (another drawing's corner) are wiped before rendering.

The figure is then rendered from a private copy of the page from which the
foreign text has been removed (a real removal, not a white patch), as a PNG and,
for line drawings, optionally as an SVG that stays sharp at any zoom.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import fitz  # PyMuPDF

try:
    import cv2
    import numpy as np
except Exception:  # pragma: no cover - without OpenCV the model's own box is used
    cv2 = None
    np = None

MARGIN = 30.0     # how far beyond the guess a figure may reach, points
JOIN = 2.5        # ink closer than this belongs to the same piece, points

ANALYSIS_DPI = 110
INSIDE = 0.4      # share of a piece's ink that must lie inside the guess


@dataclass
class Cut:
    rect: fitz.Rect                                  # the refined crop, page points
    foreign: list = field(default_factory=list)      # word boxes to remove (text layer)
    wipes: list = field(default_factory=list)        # areas to blank (neighbour pieces)
    vector: bool = False                             # a pure line drawing (no bitmap in it)
    moved: float = 0.0                               # largest edge shift against the guess


def _area(r: fitz.Rect) -> float:
    return max(0.0, r.width) * max(0.0, r.height)


def _share_inside(inner: fitz.Rect, outer: fitz.Rect) -> float:
    a = _area(inner)
    return _area(inner & outer) / a if a else 0.0


class FigureCutter:
    def __init__(self, doc: fitz.Document):
        self.doc = doc
        self._words: dict[int, list] = {}
        self._images: dict[int, list] = {}

    # ------------------------------------------------------------------ #
    def words(self, pno: int) -> list:
        if pno not in self._words:
            self._words[pno] = [fitz.Rect(w[:4]) for w in self.doc[pno].get_text("words")]
        return self._words[pno]

    def images(self, pno: int) -> list:
        if pno not in self._images:
            page = self.doc[pno]
            self._images[pno] = [fitz.Rect(i["bbox"]) for i in page.get_image_info()
                                 if _area(fitz.Rect(i["bbox"])) < 0.6 * _area(page.rect)]
        return self._images[pno]

    def foreign_words(self, pno: int, seed: fitz.Rect, claimed: list) -> list:
        """Words of the text blocks that are not part of the figure.

        A block lying mostly inside the guess is the figure's own lettering (the
        layout model often reads a diagram's labels as a text block of its own).
        """
        outside = [b for b in claimed if _share_inside(b, seed) < 0.5]
        out = []
        for w in self.words(pno):
            c = fitz.Point((w.x0 + w.x1) / 2, (w.y0 + w.y1) / 2)
            if any(b.contains(c) for b in outside):
                out.append(w)
        return out

    # ------------------------------------------------------------------ #
    def refine(self, pno: int, seed: fitz.Rect, others: list, claimed: list,
               solid: list | None = None) -> Cut:
        """The figure's real box near ``seed``.

        ``others``: the guesses for the other figures of the page.
        ``claimed``: boxes of the text / table blocks around (never part of a figure).
        ``solid``: areas that are foreign as a whole - a table with its borders, a
        caption that is printed as text again under the figure.
        """
        page = self.doc[pno]
        seed = fitz.Rect(seed) & page.rect
        has_text = bool(self.words(pno))
        foreign = self.foreign_words(pno, seed, claimed) if has_text else []
        cut = Cut(fitz.Rect(seed), foreign)
        cut.vector = not any(_area(i & seed) > 0.2 * min(_area(i), _area(seed))
                             for i in self.images(pno)) and bool(page.get_drawings())
        if cv2 is None or seed.width < 4 or seed.height < 4:
            return cut

        win = (seed + (-MARGIN, -MARGIN, MARGIN, MARGIN)) & page.rect
        pm = page.get_pixmap(clip=win, dpi=ANALYSIS_DPI, colorspace=fitz.csGRAY, alpha=False)
        img = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.stride)[:, :pm.width]
        ink = img < 215
        sx, sy = pm.width / win.width, pm.height / win.height

        def box(r: fitz.Rect, grow: float = 0.0) -> tuple[int, int, int, int]:
            x0 = int(max(0, (r.x0 - grow - win.x0) * sx))
            y0 = int(max(0, (r.y0 - grow - win.y0) * sy))
            x1 = int(min(pm.width, (r.x1 + grow - win.x0) * sx + 1))
            y1 = int(min(pm.height, (r.y1 + grow - win.y0) * sy + 1))
            return x0, y0, max(x0, x1), max(y0, y1)

        ink = ink.copy()
        # 1. take away what is not the figure
        blank = foreign if has_text else [b for b in claimed if _share_inside(b, seed) < 0.5]
        if has_text:
            # text the page holds only as a picture (no words in the text layer) but
            # the layout model read as a text block: foreign as an area
            centres = [fitz.Point((w.x0 + w.x1) / 2, (w.y0 + w.y1) / 2) for w in self.words(pno)]
            pixel_text = [b for b in claimed if _share_inside(b, seed) < 0.5
                          and b.intersects(seed + (-MARGIN, -MARGIN, MARGIN, MARGIN))
                          and not any(b.contains(c) for c in centres)]
            solid = list(solid or []) + pixel_text
        # a table / caption beside the figure goes as a whole (unless the guess is so
        # far off that it lies well inside it - then it cannot be told apart)
        solid = [s + (-1.5, -1.5, 1.5, 1.5) for s in (solid or [])
                 if _share_inside(seed, s) < 0.15 and _share_inside(s, seed) < 0.5
                 or _share_inside(s, seed) >= 0.5 and _area(s) < 0.25 * _area(seed)]
        for r in blank:
            x0, y0, x1, y1 = box(r, 0.6)
            ink[y0:y1, x0:x1] = False
        for r in solid:
            x0, y0, x1, y1 = box(r)
            ink[y0:y1, x0:x1] = False
        sx0, sy0, sx1, sy1 = box(seed)
        own = np.zeros_like(ink)
        own[sy0:sy1, sx0:sx1] = True
        theirs = np.zeros_like(ink)  # the other figures' ground (outside this guess)
        for o in others:
            x0, y0, x1, y1 = box(o)
            theirs[y0:y1, x0:x1] = True
        theirs &= ~own
        # rules: a column separator / a table edge running through the whole window
        n, lab, stats, _ = cv2.connectedComponentsWithStats(ink.astype(np.uint8), connectivity=8)
        for k in range(1, n):
            x, y, w, h, _a = stats[k]
            if (w <= 4 and h >= 0.85 * pm.height) or (h <= 4 and w >= 0.85 * pm.width):
                ink[lab == k] = False
        if not ink[sy0:sy1, sx0:sx1].any():
            return cut

        # 2. pieces of ink (strokes that touch) and groups of pieces lying close
        #    together; a group mostly inside the guess is the figure
        n, lab, stats, _ = cv2.connectedComponentsWithStats(ink.astype(np.uint8), connectivity=8)
        k = max(1, int(round(JOIN / 2 * ANALYSIS_DPI / 72)))  # both sides grow: gap <= JOIN
        grown = cv2.dilate(ink.astype(np.uint8), np.ones((2 * k + 1, 2 * k + 1), np.uint8))
        g_n, g_lab = cv2.connectedComponents(grown, connectivity=8)
        total = np.bincount(lab.ravel(), minlength=n).astype(float)
        mine_px = np.bincount(lab[own].ravel(), minlength=n).astype(float)
        their_px = np.bincount(lab[theirs].ravel(), minlength=n).astype(float)
        group = np.zeros(n, dtype=int)  # the group every piece belongs to
        group[lab[ink]] = g_lab[ink]
        g_total = np.bincount(group, weights=total, minlength=g_n)
        g_mine = np.bincount(group, weights=mine_px, minlength=g_n)
        g_their = np.bincount(group, weights=their_px, minlength=g_n)
        mine = set()
        for i in range(1, n):
            g = group[i]
            if not total[i] or not g_total[g]:
                continue
            if g_their[g] < 0.15 * g_total[g]:      # nobody else claims the group
                if g_mine[g] >= INSIDE * g_total[g]:
                    mine.add(i)
            elif mine_px[i] > their_px[i] and mine_px[i] >= INSIDE * total[i]:
                mine.add(i)                         # two figures touch: piece by piece
        if not mine:
            return cut

        def extent(ids) -> tuple[int, int, int, int]:
            ys, xs = np.where(np.isin(lab, list(ids)))
            return xs.min(), ys.min(), xs.max() + 1, ys.max() + 1

        x0, y0, x1, y1 = extent(mine)
        # 3. small loose pieces right beside it (an axis letter, a "t", a charge sign),
        #    unless they stand nearer to another figure
        reach = int(8 * ANALYSIS_DPI / 72)
        other_boxes = [box(o) for o in others]

        def gap(a, b) -> float:  # distance between two pixel boxes
            return max(a[0] - b[2], b[0] - a[2], a[1] - b[3], b[1] - a[3], 0)

        taken = {group[i] for i in mine}
        for g in range(1, g_n):  # whole groups only: one letter of a text line is no label
            if g in taken or not g_total[g] or g_their[g] > g_mine[g]:
                continue
            ids = [i for i in range(1, n) if group[i] == g and total[i]]
            gx0 = min(stats[i][0] for i in ids)
            gy0 = min(stats[i][1] for i in ids)
            gx1 = max(stats[i][0] + stats[i][2] for i in ids)
            gy1 = max(stats[i][1] + stats[i][3] for i in ids)
            me = (gx0, gy0, gx1, gy1)
            if gx1 - gx0 <= 45 * sx and gy1 - gy0 <= 16 * sy \
                    and gap(me, (x0, y0, x1, y1)) <= reach \
                    and all(gap(me, o) > gap(me, (x0, y0, x1, y1)) for o in other_boxes):
                mine.update(ids)
                continue
            # a name drawn (not typed) right under the figure is its caption
            under = -2 * sy <= gy0 - y1 <= 6 * sy and gx0 < x1 and gx1 > x0
            area = fitz.Rect(win.x0 + gx0 / sx, win.y0 + gy0 / sy,
                             win.x0 + gx1 / sx, win.y0 + gy1 / sy)
            if under and gy1 - gy0 <= 16 * sy and gx1 - gx0 <= 200 * sx \
                    and g_their[g] < 0.6 * g_total[g] \
                    and not any(_area(w & area) > 0 for w in self.words(pno)):
                mine.update(ids)


        x0, y0, x1, y1 = extent(mine)
        rect = fitz.Rect(win.x0 + x0 / sx, win.y0 + y0 / sy, win.x0 + x1 / sx, win.y0 + y1 / sy)
        rect = (rect + (-1.2, -1.2, 1.2, 1.2)) & page.rect
        # 4. neighbours that still fall inside the box
        member = np.isin(lab, list(mine))
        for i in range(1, n):
            if i in mine or not total[i]:
                continue
            x, y, w, h, _a = stats[i]
            r = fitz.Rect(win.x0 + x / sx - 0.8, win.y0 + y / sy - 0.8,
                          win.x0 + (x + w) / sx + 0.8, win.y0 + (y + h) / sy + 0.8) & rect
            if _area(r) > 0:
                bx0, by0, bx1, by1 = box(r)
                if not member[by0:by1, bx0:bx1].any():
                    cut.wipes.append(r)
        if not has_text:  # scanned page: text is blanked as areas
            cut.wipes += [b & rect for b in blank if _area(b & rect) > 0]
        cut.wipes += [s & rect for s in solid if _area(s & rect) > 0]

        cut.moved = max(abs(rect.x0 - seed.x0), abs(rect.y0 - seed.y0),
                        abs(rect.x1 - seed.x1), abs(rect.y1 - seed.y1))
        cut.rect = rect
        return cut

    # ------------------------------------------------------------------ #
    def render(self, pno: int, cut: Cut, out: Path, dpi: int, svg: Path | None = None) -> None:
        """PNG of the crop (and an SVG when asked) without the foreign text / pieces."""
        tmp = fitz.open()
        tmp.insert_pdf(self.doc, from_page=pno, to_page=pno)
        page = tmp[0]
        if page.rotation:
            page.remove_rotation()
        hit = [w for w in cut.foreign if not (w & cut.rect).is_empty]
        if hit:
            for w in hit:
                page.add_redact_annot(w + (0.3, 0.6, -0.3, -0.6), fill=False)
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                  graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        for r in cut.wipes:
            page.draw_rect(r, color=None, fill=(1, 1, 1), overlay=True)
        page.get_pixmap(clip=cut.rect, dpi=dpi, alpha=False).save(out)
        if svg is not None:
            text = vector_svg(page, cut)
            if text is not None:
                svg.write_text(text, encoding="utf-8")
            elif svg.exists():
                svg.unlink()
        tmp.close()


def _rgb(c) -> str:
    return "#%02x%02x%02x" % tuple(int(round(max(0.0, min(1.0, v)) * 255)) for v in c[:3])


def vector_svg(page: fitz.Page, cut: Cut) -> str | None:
    """The figure's own strokes as a small SVG, or None when that would not be exact.

    Only the paths of the drawing itself are written (no page content around it).
    The result is rendered and compared with the page: a drawing that uses
    something this writer does not carry over (a clip, a shading, real text) does
    not match and stays a PNG.
    """
    r = cut.rect
    if any(not (fitz.Rect(w[:4]) & r).is_empty for w in page.get_text("words", clip=r)):
        return None  # typed lettering inside: keep the exact picture
    out = []
    for d in page.get_drawings():
        box = fitz.Rect(d["rect"])
        c = fitz.Point((box.x0 + box.x1) / 2, (box.y0 + box.y1) / 2)
        if not (r + (-0.5, -0.5, 0.5, 0.5)).contains(c) or any(w.contains(c) for w in cut.wipes):
            continue
        path, pen = [], None
        for it in d["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if pen != a:
                    path.append(f"M{a.x - r.x0:.2f} {a.y - r.y0:.2f}")
                path.append(f"L{b.x - r.x0:.2f} {b.y - r.y0:.2f}")
                pen = b
            elif it[0] == "c":
                a, b1, b2, b = it[1:5]
                if pen != a:
                    path.append(f"M{a.x - r.x0:.2f} {a.y - r.y0:.2f}")
                path.append(f"C{b1.x - r.x0:.2f} {b1.y - r.y0:.2f} {b2.x - r.x0:.2f} "
                            f"{b2.y - r.y0:.2f} {b.x - r.x0:.2f} {b.y - r.y0:.2f}")
                pen = b
            elif it[0] == "re":
                q = it[1]
                path.append(f"M{q.x0 - r.x0:.2f} {q.y0 - r.y0:.2f}H{q.x1 - r.x0:.2f}"
                            f"V{q.y1 - r.y0:.2f}H{q.x0 - r.x0:.2f}Z")
                pen = None
            elif it[0] == "qu":
                q = it[1]
                pts = (q.ul, q.ur, q.lr, q.ll)
                path.append("M" + "L".join(f"{p.x - r.x0:.2f} {p.y - r.y0:.2f}" for p in pts) + "Z")
                pen = None
        if not path:
            continue
        if d.get("closePath"):
            path.append("Z")
        fill = d.get("fill") if "f" in d["type"] else None
        stroke = d.get("color") if "s" in d["type"] else None
        attr = [f'd="{"".join(path)}"', f'fill="{_rgb(fill) if fill else "none"}"']
        if fill and d.get("even_odd"):
            attr.append('fill-rule="evenodd"')
        if fill and (d.get("fill_opacity") or 1) < 1:
            attr.append(f'fill-opacity="{d["fill_opacity"]:.2f}"')
        if stroke:
            attr.append(f'stroke="{_rgb(stroke)}" stroke-width="{max(d.get("width") or 0, 0.2):.2f}"')
            cap = (d.get("lineCap") or (0,))[0]
            attr.append('stroke-linecap="%s"' % ("butt", "round", "square")[min(cap, 2)])
            attr.append('stroke-linejoin="%s"' % ("miter", "round", "bevel")[
                min(int(d.get("lineJoin") or 0), 2)])
            dash = (d.get("dashes") or "").strip()
            nums = dash[dash.find("[") + 1:dash.find("]")].split() if "[" in dash else []
            if nums:
                attr.append(f'stroke-dasharray="{" ".join(nums)}"')
            if (d.get("stroke_opacity") or 1) < 1:
                attr.append(f'stroke-opacity="{d["stroke_opacity"]:.2f}"')
        out.append("<path " + " ".join(attr) + "/>")
    if not out:
        return None
    text = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{r.width:.2f}pt" '
            f'height="{r.height:.2f}pt" viewBox="0 0 {r.width:.2f} {r.height:.2f}">'
            + "".join(out) + "</svg>")
    return text if _same_picture(page, r, text) else None


def _same_picture(page: fitz.Page, r: fitz.Rect, svg: str) -> bool:
    """Does the SVG draw what the page shows in ``r``? (ink compared at ~100 dpi)"""
    if cv2 is None:
        return False
    try:
        with fitz.open(stream=svg.encode("utf-8"), filetype="svg") as d:
            a = d[0].get_pixmap(dpi=100, colorspace=fitz.csGRAY, alpha=False)
    except Exception:
        return False
    b = page.get_pixmap(clip=r, dpi=100, colorspace=fitz.csGRAY, alpha=False)

    def ink(pm):
        return np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.stride)[:, :pm.width] < 170

    h, w = min(a.height, b.height), min(a.width, b.width)
    if abs(a.height - b.height) > 2 or abs(a.width - b.width) > 2 or not h or not w:
        return False
    x, y = ink(a)[:h, :w], ink(b)[:h, :w]
    near = np.ones((3, 3), np.uint8)
    miss = (x & ~cv2.dilate(y.astype(np.uint8), near).astype(bool)).sum() \
        + (y & ~cv2.dilate(x.astype(np.uint8), near).astype(bool)).sum()
    return miss <= 0.02 * max(int(y.sum()), 1)

