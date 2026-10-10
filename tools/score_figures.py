"""Score figure crops of a converted paper against the PDF itself.

  python score_figures.py PAPER.pdf CACHE_DIR            (CACHE_DIR = <name>_mineru_basic)
  python score_figures.py --bench FOLDER

For every figure the layout model found, two crops are judged - the model's own
box (what the converter used before) and the refined box (what it uses now):

  cut      ink of the figure touches the crop's edge: a stroke or a letter is cut
  foreign  words of the surrounding text (PDF text layer) fall inside the crop
  loose    blank margin around the figure, points (tightness)

No reference file is needed; the page is its own ground truth.
"""
import json
import sys
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parent))
from figure_cut import FigureCutter, _share_inside  # noqa: E402

try:
    import numpy as np
except ImportError:  # pragma: no cover
    sys.exit("numpy is needed for the crop check")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
TEXT = {"text", "title", "paragraph_title", "doc_title", "list", "index", "table"}


def judge(page: fitz.Page, crop: fitz.Rect, foreign: list, wipes: list) -> tuple[bool, int, float]:
    """(cut, foreign words inside, blank margin) for one crop of the page."""
    crop = crop & page.rect
    pm = page.get_pixmap(clip=crop, dpi=144, colorspace=fitz.csGRAY, alpha=False)
    a = np.frombuffer(pm.samples, dtype=np.uint8).reshape(pm.height, pm.stride)[:, :pm.width] < 200
    a = a.copy()
    s = pm.width / max(crop.width, 1e-6)
    for r in list(foreign) + list(wipes):  # what the renderer removes is not ink of the figure
        x0, y0 = int(max(0, (r.x0 - crop.x0) * s)), int(max(0, (r.y0 - crop.y0) * s))
        x1, y1 = int(max(0, (r.x1 - crop.x0) * s + 1)), int(max(0, (r.y1 - crop.y0) * s + 1))
        a[y0:y1, x0:x1] = False
    edge = int(a[0].sum() + a[-1].sum() + a[:, 0].sum() + a[:, -1].sum())
    inside = sum(1 for w in foreign if _share_inside(w, crop) > 0.3)
    if not a.any():
        return False, inside, 0.0
    ys, xs = np.where(a)
    margin = min(xs.min(), ys.min(), pm.width - 1 - xs.max(), pm.height - 1 - ys.max()) / s
    return edge > 2, inside, float(margin)


def score(pdf: Path, cache: Path, verbose: bool = False) -> dict:
    doc = fitz.open(pdf)
    m = json.loads((cache / "middle_json.json").read_text(encoding="utf-8"))
    cutter = FigureCutter(doc)
    out = {"figures": 0, "old_cut": 0, "new_cut": 0, "old_foreign": 0, "new_foreign": 0,
           "moved": 0}
    for page in m["pages"]:
        pno = page["page_idx"]
        p = doc[pno]
        W, H = p.rect.width, p.rect.height
        rect = lambda bb: fitz.Rect(bb[0] * W, bb[1] * H, bb[2] * W, bb[3] * H)  # noqa: E731
        seeds, claimed, solid = [], [], []
        for b in page["blocks"]:
            parts = [q for q in (b.get("content") or []) if isinstance(q, dict) and q.get("bbox")] \
                if isinstance(b.get("content"), list) else []
            if b["type"] == "image":
                seeds += [rect(q["bbox"]) for q in parts if q.get("type") == "image_body"] \
                    or [rect(b["bbox"])]
                caps = [rect(q["bbox"]) for q in parts if q.get("type") != "image_body"]
                claimed += caps
                solid += caps
            elif b["type"] in TEXT:
                r = rect(b["bbox"])
                if any(_share_inside(r, s_) > 0.7 for s_ in
                       [rect(x["bbox"]) for x in page["blocks"] if x["type"] == "image"]):
                    continue  # the figure's own lettering read as text
                claimed.append(r)
                if b["type"] == "table":
                    solid.append(r)
        for seed in seeds:
            words = cutter.foreign_words(pno, seed, claimed)
            old = (seed + (-2, -2, 2, 2)) & p.rect
            cut = cutter.refine(pno, seed, [o for o in seeds if o != seed], claimed, solid)
            oc, of, _ = judge(p, old, [], [])
            of = sum(1 for w in words if _share_inside(w, old) > 0.3)
            nc, nf, _ = judge(p, cut.rect, cut.foreign, cut.wipes)
            nf = 0 if cut.foreign is not None else nf  # removed before rendering
            out["figures"] += 1
            out["old_cut"] += oc
            out["new_cut"] += nc
            out["old_foreign"] += of > 0
            out["new_foreign"] += nf > 0
            out["moved"] += cut.moved > 3
            if verbose and (oc or of or nc):
                print(f"  p{pno + 1} {[round(v) for v in seed]}: old cut={oc} foreign={of}"
                      f" -> new {[round(v) for v in cut.rect]} cut={nc}")
    return out


def report(name: str, r: dict) -> None:
    n = r["figures"] or 1
    print(f"{name:10s} figures {r['figures']:3d} | cut at the edge: {r['old_cut']:2d} -> "
          f"{r['new_cut']:2d} | neighbouring text inside: {r['old_foreign']:2d} -> "
          f"{r['new_foreign']:2d} | boxes moved >3pt: {r['moved']:2d} ({100 * r['moved'] // n}%)")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "-v"]
    verbose = "-v" in sys.argv
    if args and args[0] == "--bench":
        folder = Path(args[1])
        for pdf in sorted(folder.glob("*.pdf")):
            cache = folder / f"{pdf.stem}_mineru_basic"
            if (cache / "middle_json.json").exists():
                report(pdf.stem, score(pdf, cache, verbose))
    elif len(args) == 2:
        report(Path(args[0]).stem[:10], score(Path(args[0]), Path(args[1]), verbose))
    else:
        sys.exit(__doc__)
