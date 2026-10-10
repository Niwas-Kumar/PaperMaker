"""How a paper writes its question numbers and option labels.

Coaching papers differ in small, systematic ways:

    numbers   12.      *12.     Q.12     Q12.     Q.12.    Q 12)    12)
    options   (1) (2)  (A) (B)  (a) (b)  A) B)    a) b)    1) 2)    A. B.

The converter reads every paper in ONE internal form - ``12.`` and ``(A)`` - so
all the later steps (splitting stems from options, option grids, repairs from the
PDF text layer, the house-style template) work the same for every style. The
style found here is written back when the Word file is produced, so the output
looks like the source.

``detect`` decides from the paper's own lines: the numbering style is the one
whose numbers form the longest 1, 2, 3 ... run (``12.`` wins a tie, so "1)" sub-
points inside a solution never change a normal paper); bare labels ``A)`` / ``A.``
are only taken as options when the paper clearly uses them and not ``(A)``.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

Q_WORD = r"(?:Q|Que|Ques|Question)"
DOT_RE = re.compile(r"^\s*\*?\s*(\d{1,3})\.(?!\d)")
Q_RE = re.compile(rf"^\s*({Q_WORD}\s*[.\-:]?\s*)(\d{{1,3}})(\s*[.):]?)(?=\s|$|[A-Z])")
PAREN_RE = re.compile(r"^\s*(\d{1,3})\)(?=\s|$|[A-Z])")

MASK_RE = re.compile(r"\$\$.+?\$\$|\$[^$\n]+?\$|<[^>]+>|⟪[^⟫]*⟫", re.S)
LEAD_RE = re.compile(r"^((?:\s|<[^>]+>)*)(.*)$", re.S)


@dataclass
class PaperFormat:
    number: str = "{n}."       # the source's numbering: "Q.{n}", "Q{n}.", "{n})"
    label: str = "({x})"       # the source's option labels: "{x})", "{x}."
    chars: str = ""            # which bare labels are options: "A-E", "a-e", "1-5"

    @property
    def plain(self) -> bool:
        return self.number == "{n}." and self.label == "({x})"

    def describe(self) -> str:
        return (f"numbers like {self.number.format(n=12)}, options like "
                f"{self.label.format(x=(self.chars or 'A')[0])}")

    # -- reading: the source's style -> the internal form ------------------ #
    def _number_match(self, text: str):
        if self.number.endswith("{n})") and not self.number.startswith("Q"):
            m = PAREN_RE.match(text)
            return (m, m.group(1)) if m else (None, None)
        m = Q_RE.match(text)
        return (m, m.group(2)) if m else (None, None)

    def line(self, s: str) -> str:
        """One converted line (HTML with $maths$) in the internal form."""
        if self.plain or not s:
            return s
        if self.number != "{n}.":
            head, rest = LEAD_RE.match(s).groups()
            m, n = self._number_match(rest)
            if m:
                s = f"{head}{n}. {rest[m.end():].lstrip()}"
        if self.label != "({x})":
            close = re.escape(self.label[-1])
            rx = re.compile(rf"(?:(?<=\s)|^)([{self.chars}]){close}(?=\s|$)")
            parts, keep = MASK_RE.split(s), MASK_RE.findall(s)
            parts = [rx.sub(r"(\1)", p) for p in parts]
            s = "".join(x for pair in zip(parts, keep + [""]) for x in pair)
        return s

    def split(self, s: str) -> list[str]:
        """"... (D) text Q.106 Match List-I": the next question glued to the end of a
        line starts its own line (only the unmistakable "Q.12" style is split)."""
        if not self.number.startswith("Q") or s.lstrip().startswith("<table"):
            return [s]
        out, rest = [], s
        while True:
            m = re.search(rf"(?<=\S)\s+(?={Q_WORD}\s*[.\-:]?\s*\d{{1,3}}\s*[.):]?(?:\s|$|[A-Z]))",
                          MASK_RE.sub(lambda k: "\0" * len(k.group()), rest))
            if not m:
                break
            out.append(rest[:m.start()])
            rest = self.line(rest[m.end():])
        return out + [rest]

    def words(self, words: list) -> list:

        """PyMuPDF words with numbers / labels in the internal form."""
        if self.plain:
            return words
        out, skip = [], False
        q_only = re.compile(rf"{Q_WORD}[.\-:]?")
        close = re.escape(self.label[-1])
        lab = re.compile(rf"([{self.chars}]){close}") if self.label != "({x})" else None
        for i, w in enumerate(words):
            if skip:
                skip = False
                continue
            t = w[4]
            if self.number != "{n}.":
                nxt = words[i + 1] if i + 1 < len(words) else None
                if q_only.fullmatch(t) and nxt is not None and abs(nxt[1] - w[1]) < 3 \
                        and re.fullmatch(r"(\d{1,3})[.):]?", nxt[4]):  # "Q." + "12"
                    n = re.match(r"\d+", nxt[4]).group()
                    out.append((w[0], min(w[1], nxt[1]), nxt[2], max(w[3], nxt[3]), f"{n}.")
                               + tuple(w[5:]))
                    skip = True
                    continue
                m, n = self._number_match(t)
                if m and m.end() == len(t):
                    t = f"{n}."
            if lab is not None:
                m = lab.fullmatch(t)
                if m:
                    t = f"({m.group(1)})"
            out.append(tuple(w[:4]) + (t,) + tuple(w[5:]))
        return out

    # -- writing: the internal form -> the source's style ------------------ #
    def number_text(self, digits: str, star: str = "") -> str:
        return star + self.number.format(n=digits)

    def label_text(self, text: str) -> str:
        """Option labels of one run of an option row: "(A) x<tab>(B) y" -> "A) x<tab>B) y"."""
        if self.label == "({x})":
            return text
        return re.sub(rf"(?:(?<=\t)|^)\(([{self.chars}])\)(?=\s|$)",
                      lambda m: self.label.format(x=m.group(1)), text)


def detect(lines: list[str]) -> PaperFormat:
    """The paper's style, from its lines (PDF text layer, or the OCR text of a scan)."""
    lines = [ln.strip() for ln in lines if ln and ln.strip()]
    fmt = PaperFormat()

    def run_of(rx, group: int) -> int:
        longest = run = 0
        last = None
        for ln in lines:
            m = rx.match(ln)
            if not m:
                continue
            n = int(m.group(group))
            run = run + 1 if last is not None and n == last + 1 else 1
            longest, last = max(longest, run), n
        return longest

    dot, q, paren = run_of(DOT_RE, 1), run_of(Q_RE, 2), run_of(PAREN_RE, 1)
    if q >= 5 and q > dot:
        shapes = Counter((re.sub(r"\s+", " ", m.group(1)), m.group(3).strip())
                         for ln in lines if (m := Q_RE.match(ln)))
        (prefix, suffix), _ = shapes.most_common(1)[0]
        fmt.number = prefix.strip() + ("" if prefix.strip()[-1:] in ".-:" else
                                       " " if prefix.endswith(" ") else "") + "{n}" + suffix
    elif paren >= 5 and paren > dot and paren > q:
        fmt.number = "{n})"

    text = "\n".join(lines)

    def count(pattern: str) -> Counter:
        return Counter(m.group(1) for m in re.finditer(pattern, text, re.M))

    bracket = sum(count(r"(?:(?<=\s)|^)\(([A-Da-d1-4])\)(?=\s)").values())
    for chars, klass in (("A-E", "A-D"), ("a-e", "a-d"), ("1-5", "1-4")):
        if chars == "1-5" and fmt.number == "{n})":
            continue  # "1)" is this paper's question number
        bare = count(rf"(?:(?<=\s)|^)([{klass}])\)(?=\s)")
        if len(bare) == 4 and min(bare.values()) >= 5 and sum(bare.values()) > 2 * bracket:
            fmt.label, fmt.chars = "{x})", chars
            return fmt
    if bracket < 8:  # "A. text  B. text": only when the paper has no bracketed labels
        for chars, klass in (("A-E", "A-D"), ("a-e", "a-d")):
            dotted = count(rf"(?:(?<=\s)|^)([{klass}])\.(?=\s)")
            if len(dotted) == 4 and min(dotted.values()) >= 8 \
                    and max(dotted.values()) <= 1.5 * min(dotted.values()):
                fmt.label, fmt.chars = "{x}.", chars
                break
    return fmt


FORMAT = PaperFormat()  # the paper being converted (set once by the converter)
