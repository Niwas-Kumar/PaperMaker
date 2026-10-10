"""Score a converted .docx against the original Word paper, question by question.

  python score_conversion.py ORIGINAL.docx CONVERTED.docx [-v]
  python score_conversion.py --bench FOLDER

--bench re-converts every X.pdf in FOLDER that has an X.docx beside it (MinerU
results are cached there, so this takes seconds) and scores X_out.docx against
X.docx. Run it after every change to pdf_math_to_docx.py / house_style.py.

Checks per question: found, option count, 2-up/1-up rows, A./I./bullet items,
tables, empty option cells, text similarity; plus the answer key and solution
headings. Formulas that are MathType objects in the original have no text, so a
low text score on a formula-heavy question is expected.
"""
import difflib
import re
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from house_style import question_start  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
QN = re.compile(r"^\s*\*?\s*(\d{1,3})\.\s*\t")
KEY = re.compile(r"^\s*(\d{1,3})\.\s*\t?\s*\(\s*([1-5a-dA-D])\s*\)\s*$")
SOL = re.compile(r"^\s*(\d{1,3})\.\s*\t?\s*\(\s*([1-5a-dA-D])\s*\)")
OPT = re.compile(r"\(\s*(?:[1-5]|[a-e])\s*\)")


def ft(p):
    out = []
    for el in p.iter(qn("w:t"), qn("w:tab"), qn("m:t")):
        if el.tag == qn("w:tab"):
            if el.getparent().tag == qn("w:r"):
                out.append("\t")
        else:
            out.append(el.text or "")
    return "".join(out)


def norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


def parse(path):
    body = Document(path).element.body
    q, key, sol = {}, {}, {}
    phase, cur = "pre", None
    kids = list(body.iterchildren())
    numbered = [(i, int(m.group(1))) for i, el in enumerate(kids)
                if el.tag == qn("w:p") and (m := QN.match(ft(el)))]
    begin = numbered[question_start([n for _, n in numbered])][0] if numbered else 0
    for i, el in enumerate(kids):
        if i < begin:  # cover, incl. numbered instructions
            continue
        if el.tag == qn("w:tbl"):
            if cur is not None and phase == "q":
                q[cur]["tables"] += 1
                rows = el.findall(qn("w:tr"))
                q[cur]["text"] += " " + " ".join(ft(p) for p in el.iter(qn("w:p")))
                q[cur]["rows"] += len(rows)
            continue
        if el.tag != qn("w:p"):
            continue
        t = ft(el)
        s = t.strip()
        if re.fullmatch(r"ANSWER\s*KEY", s, re.I):
            phase = "key"
            continue
        if phase == "key" and SOL.match(t) and not KEY.match(t):
            phase = "sol"
        m = QN.match(t)
        if phase in ("pre", "q") and m and not KEY.match(t):
            n = int(m.group(1))
            if phase == "pre" or (cur is not None and n > cur):
                phase, cur = "q", n
                q[n] = dict(empty=0, opts=0, rows2=0, rows1=0, items=0, tables=0, rows=0,
                            pics=0, text=t[m.end():], paras=1)
                q[n]["pics"] += len(list(el.iter(qn("w:drawing"))))
                continue
        if phase == "q" and cur is not None:
            d = q[cur]
            d["pics"] += len(list(el.iter(qn("w:drawing")))) + len(list(el.iter(qn("w:object"))))
            if not s:
                continue
            labs = OPT.findall(t)
            if re.match(r"^\t?\(\s*(?:[1-5]|[a-e])\s*\)\s*\t", t):  # "(1)<tab>" or "(1) <tab>"
                d["opts"] += len(labs)
                has_obj = el.find(".//" + qn("w:drawing")) is not None or \
                    el.find(".//" + qn("w:object")) is not None
                if not has_obj and re.search(r"\(\s*[1-5a-e]\s*\)\s*\t\s*(?:\t|$)", t):
                    d["empty"] += 1
                d["rows2" if len(labs) >= 2 else "rows1"] += 1
            elif re.match(r"^\t(?:[A-E]|I{1,3}|IV|V|•)\.?\t", t):
                d["items"] += 1
                d["text"] += " " + t
            else:
                d["paras"] += 1
                d["text"] += " " + t
        elif phase in ("key", "sol"):
            for a, b in re.findall(r"(\d{1,3})\.\s*\t?\s*\(\s*([1-5a-dA-D])\s*\)", t):
                (key if phase == "key" else sol).setdefault(int(a), b)
    return q, key, sol


def bench(folder: Path) -> None:
    tool = Path(__file__).with_name("pdf_math_to_docx.py")
    for pdf in sorted(folder.glob("*.pdf")):
        orig = pdf.with_suffix(".docx")
        if not orig.exists():
            continue
        out = pdf.with_name(pdf.stem + "_out.docx")
        print(f"########## {pdf.stem}")
        r = subprocess.run([sys.executable, str(tool), str(pdf), "-o", str(out)],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode:
            print(r.stdout[-1500:], r.stderr[-1500:])
            continue
        score(orig, out, False)


def main():
    if sys.argv[1] == "--bench":
        bench(Path(sys.argv[2]))
    else:
        score(sys.argv[1], sys.argv[2], "-v" in sys.argv)


def score(orig, conv, verbose):
    q0, k0, s0 = parse(orig)
    q1, k1, s1 = parse(conv)
    nums = sorted(q0)
    found = [n for n in nums if n in q1]
    extra = sorted(set(q1) - set(q0))
    bad = {k: [] for k in ("opts", "layout", "items", "tables", "text")}
    sims = []
    for n in found:
        a, b = q0[n], q1[n]
        if a["opts"] != b["opts"]:
            bad["opts"].append(f"{n}:{a['opts']}->{b['opts']}")
        elif (a["rows2"], a["rows1"]) != (b["rows2"], b["rows1"]):
            bad["layout"].append(f"{n}:{a['rows2']}x2+{a['rows1']}x1->{b['rows2']}x2+{b['rows1']}x1")
        if a["items"] != b["items"]:
            bad["items"].append(f"{n}:{a['items']}->{b['items']}")
        if a["tables"] != b["tables"]:
            bad["tables"].append(f"{n}:{a['tables']}->{b['tables']}")
        sim = difflib.SequenceMatcher(None, norm(a["text"]), norm(b["text"])).ratio()
        sims.append(sim)
        if sim < 0.85:
            bad["text"].append(f"{n}:{sim:.2f}")
    print(f"questions   {len(found)}/{len(nums)} found" + (f", extra {extra}" if extra else ""))
    missing = [n for n in nums if n not in q1]
    if missing:
        print(f"   missing  {missing}")
    for k, v in bad.items():
        print(f"{k:11s} {len(found) - len(v)}/{len(found)} ok" + (f"   bad: {' '.join(v[:25])}" if v else ""))
    empty = [n for n in found if q1[n]["empty"] and not q0[n]["empty"]]
    print(f"empty opts  {len(found) - len(empty)}/{len(found)} ok" + (f"   bad: {empty}" if empty else ""))
    if sims:
        print(f"text sim    mean {sum(sims) / len(sims):.3f}  (math in the original is MathType -> no text)")
    if k0:
        ok = sum(1 for n in k0 if k1.get(n) == k0[n])
        print(f"answer key  {ok}/{len(k0)}" + (f"   wrong: {[n for n in k0 if k1.get(n) != k0[n]][:20]}" if ok < len(k0) else ""))
    if s0:
        ok = sum(1 for n in s0 if s1.get(n) == s0[n])
        print(f"solutions   {ok}/{len(s0)} headings")
    if verbose:
        for n in found:
            a, b = q0[n], q1[n]
            print(f"--- {n}\n  O: {a['text'][:160]!r}\n  C: {b['text'][:160]!r}")


if __name__ == "__main__":
    main()
