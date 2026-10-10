# Paper Maker

Paper Maker is a Windows desktop tool for selecting, reordering, and exporting
questions from educational DOCX and PDF papers.

- DOCX input keeps the source document's styles, images, equations, headers,
  footers, and section layout.
- PDF input rebuilds plain text as editable Word content. Questions containing
  complex stacked mathematics, drawings, or photos may be inserted as
  high-resolution images to preserve their appearance.
- Each loaded source can generate questions, an answer key, or both.

## Run from source

Python 3.11 or newer is recommended.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python paper_maker.py
```

## Access control

Paper Maker opens only for people whose email is in column A of *Sheet1* of the
team's access sheet. The user clicks **Sign in with Google** (Google asks only
for their email address); the app sends Google's proof of the email to the
access service, which reads the list and answers allowed yes/no. Users never
need access to the sheet. The sign-in is remembered and the list is asked again
at every start, so an email taken off the sheet is refused the next time the
tool is opened. The check needs an internet connection.

How it fits together:

- **Google Cloud project "Paper Maker"** (Google Auth Platform): a *Desktop*
  OAuth client. Its JSON is `oauth_client.json` next to `paper_maker.py`; it is
  built into the exe and kept out of Git. Audience: External, published.
- **Access service** `access_service/` (Cloudflare Worker `paper-maker-access`):
  checks the Google ID token was made for this client, then reads
  `SHEET_ID` / `SHEET_RANGE` (`wrangler.toml`) with the owner's read-only Google
  token. The list is cached for one minute. Its URL is `ACCESS_URL` in
  `access_control.py`.

Deploy or change the service (in `access_service\`). `wrangler.toml` holds the
real client ID and sheet ID and is kept out of Git; on a new PC copy
`wrangler.example.toml` to `wrangler.toml` and fill it in:

```powershell
npx wrangler login
npx wrangler deploy
python set_owner_token.py   # owner signs in once; stores the read-only token as a secret
```

If the owner's token is ever revoked (password reset, removed in the Google
account's security page), sign-ins fail with "did not answer (HTTP 503)": run
`set_owner_token.py` again. Adding or removing an email needs nothing else.

### Private files (never committed)

| File | What it is |
|---|---|
| `oauth_client.json` | Google desktop OAuth client (built into the exe) |
| `access_service/wrangler.toml` | Worker settings with the real client ID and sheet ID |
| `installer/_setup_config.py` | R2 bucket address, written by `build_setup.py` |
| `samples/` | Test papers and their outputs (company content) |

Keep a copy of the first three somewhere safe outside the repository.

## DOCX command-line utility

To keep selected questions from a DOCX file in a specific order:

```powershell
python extract_questions.py source.docx output.docx 1 5 12 18
```

## Build the Windows application

```powershell
python -m pip install -r requirements-dev.txt
pyinstaller --noconfirm PaperMaker.spec
```

The build is written to `dist\PaperMaker.exe`.

### Shareable package (with PDF → Editable Word)

The **🧮 PDF → Editable Word** button opens `convert_window.py`, which runs
`tools\pdf_math_to_docx.py` in a portable engine next to the exe. Layout of the
shared folder (`D:\PW Automations\PaperMaker_build\PaperMaker`, zipped as
`PaperMaker_share.zip`, ≈2.2 GB):

```
PaperMaker.exe          pyinstaller --distpath <folder> PaperMaker.spec
engine\                 copy of the Python 3.11 install (no site-packages, Scripts,
                        Doc, tcl, test) + mineru_env's site-packages WITHOUT torch /
                        torchvision (MinerU then runs its ONNX models on the CPU) +
                        tools\pdf_math_to_docx.py, tools\house_style.py,
                        tools\figure_cut.py, tools\paper_format.py, pdf_core.py,
                        paper_core.py
mineru_home\            copy of %USERPROFILE%\.mineru (models, ≈2 GB); the app sets
                        MINERU_HOME to it and blocks model downloads
READ ME FIRST.txt
```

After changing the converter, copy the six .py files into `engine\` again. No
installation or GPU is needed on the target PC; CPU speed is ≈20–30 s per page. The local `dist` folder is
intentionally excluded from Git because the current executable is larger than
GitHub's normal per-file limit. Attach release builds to a GitHub Release rather
than committing them to the repository.

### Online installer (what users get)

Users get only `PaperMaker_Setup.exe` (≈11 MB). It downloads the shareable
package from the Cloudflare R2 bucket `papermaker`
(its public `r2.dev` address, ≈1 GB: the GGUF VLM is
left out because only `--tier basic` is used). It installs to
`%LOCALAPPDATA%\Programs\PaperMaker` with no admin rights, resumes broken downloads,
checks every part's SHA-256, self-tests the engine and adds Desktop / Start-menu
shortcuts plus an entry in Settings > Apps (uninstall).

```powershell
# 1. pack the tested shareable folder into ≤250 MiB parts + manifest.json
python installer\pack_release.py "D:\PW Automations\PaperMaker_build\PaperMaker" -o "D:\PW Automations\PaperMaker_build\release"
# 2. build the setup exe with the bucket address built in (only needed once)
python installer\build_setup.py <bucket public URL>
```

Upload the files listed in `release\UPLOAD_THESE.txt` to the bucket root,
`manifest.json` last. For an update, pack into a new folder with
`--previous <old release>\manifest.json`: unchanged components (usually engine and
models) keep their part names and are not packed or uploaded again. Users then
click **Update Paper Maker** in the Start menu (or run the setup again), and only
changed parts are downloaded. Delete parts that no manifest uses any more from
the bucket.

## Project files

- `paper_maker.py` - desktop interface
- `access_control.py` - Google sign-in window and the access check
- `access_service/` - Cloudflare Worker that answers the access check
- `paper_core.py` - DOCX parsing and generation
- `pdf_core.py` - PDF parsing and Word generation
- `extract_questions.py` - DOCX command-line utility
- `PaperMaker.spec` - PyInstaller build definition
- `assets/paper_maker.ico` - app icon (exe, shortcuts, all windows, setup); 16–256 px
- `installer/` - online installer: `pack_release.py` (download parts),
  `setup_app.py` (PaperMaker_Setup.exe), `build_setup.py`
- `tools/pdf_math_to_docx.py` - standalone PDF → editable Word converter with
  native Word equations (see below)

## PDF to editable Word with real equations

`tools/pdf_math_to_docx.py` turns a STEM exam PDF into a `.docx` in which every
formula is a native, editable Word equation. Figures are re-cropped from the PDF
at 300 dpi, and the page size, two-column layout, question numbers and option
columns follow the source.

Recognition runs locally with MinerU by default: free, no API key, and nothing
is uploaded. Mathpix (paid cloud API) can be used with `--engine mathpix`.

One-time setup:

```powershell
python -m pip install pymupdf python-docx requests pypandoc_binary
python -m venv ..\mineru_env
..\mineru_env\Scripts\pip install torch torchvision --index-url https://download.pytorch.org/whl/cu124
..\mineru_env\Scripts\pip install mineru
```

Convert:

```powershell
python tools\pdf_math_to_docx.py "paper.pdf"                  # -> paper_editable.docx
python tools\pdf_math_to_docx.py "paper.pdf" --pages 1-3 -o test.docx
python tools\pdf_math_to_docx.py "paper.pdf" --template "any PW paper.docx"
```

**House style (`--template`).** Give it an original Word paper of the same series
(or just put the paper's own `.docx` next to the PDF; it is picked up
automatically, and never overwritten). The output is then built inside that
document: its cover lines, PW header and footer, section breaks, ANSWER KEY band,
tab stops, hanging indents, line spacing, bold `N.` numbers and `(1) … (2) …`
option grid are copied exactly. When the template is the same paper, each
question and solution also gets its own spacing, so pages break where they did,
and the 2-up / one-per-row option layout is read from the PDF itself. Use
`--no-template` for the plain layout.

When the PDF has a text layer, the converter also checks MinerU's work against
it: lines MinerU skipped, missing question/solution headings, glued words
("atmidpoint"), lost option rows (e.g. True/False grids) and dropped bullets are
restored from the PDF's own text, and the 4-up / 2-up / one-per-row option
layout is copied question by question.

**Questions are never cut in two.** In the converter's output, and in every
paper Paper Maker builds from a Word or PDF source, a question stays with its
figures, tables and options in one column of one page (Word's *keep with next* /
*keep lines together* on every paragraph of the question, table rows that cannot
split). Only a question taller than a whole column can still break.

**Exact figure crops (`tools/figure_cut.py`).** The layout model only guesses a
figure's box; the converter corrects it on the page itself. The words of the text
around the figure and the ground of the other figures are taken away, the ink
that is left is grouped into pieces, and the pieces lying mostly inside the guess
are the figure: each is taken whole (no cut strokes or letters) and the crop is
the tight box around them. Text of the neighbouring lines is removed from the
picture (a real removal on a private copy of the page, not a white patch), and a
neighbouring drawing's corner inside the box is wiped. It works the same on
scanned pages. `tools\score_figures.py --bench <folder>` (or `PAPER.pdf
CACHE_DIR`) reports, for the old box and the new one, how many crops cut ink at
their edge or contain neighbouring text; no reference file is needed.

**Vector diagrams (`--vector-figures`, the "Sharp vector diagrams" box in the
window).** Line drawings are also stored as SVG next to the PNG: Word 2019 / 365
shows the SVG (sharp at any zoom), older programs the PNG. In Word, right-click >
*Convert to Shape* turns such a diagram into editable lines. Each SVG is rendered
and compared with the page before it is used; a drawing that does not match
exactly (or holds a bitmap or typed text) stays a PNG. Lettering inside a diagram
is stored in the PDF as drawn outlines, so it becomes shapes, not typed text.

**Numbering and option styles (`tools/paper_format.py`).** A paper may write
`12.`, `*12.`, `Q.12`, `Q12.`, `Q.12.` or `12)` and `(1)`, `(A)`, `(a)`, `A)`,
`a)`, `1)` or `A.`. The style is detected from the paper's own lines (the
numbering whose numbers form the longest 1, 2, 3 ... run; bare `A)` / `A.` only
when the paper clearly uses them), the paper is read in one internal form, and
the plain layout writes the original style back. The Word-to-Word side
(`paper_core.detect_number_style`) accepts the same numberings and renumbers a
`Q.12` paper as `Q.1, Q.2, ...`. Questions without options (numerical / integer
type) and section headings (`SECTION-A`, `PART-I`) need no setting.

**Figures.** A figure goes where the PDF prints it: centred in the column or as

far in from the column edge. Option figures keep their label and their grid
(`(A) [fig]  (B) [fig]` rows, 2-up or one per row as in the source), with the
label beside the middle of the picture; figures printed side by side stay side by
side with their captions under them. A Match table with drawings inside its cells
(structures, graphs) is placed as a picture of the table, so it looks exactly like
the source; its options below remain text.

Paper patterns it handles: numbering that starts anywhere (46, 91, 136 ...) and
starred numbers (`*9.`); questions-only papers and papers with ANSWER KEY +
solutions; options labelled `(1)..(4)`, `(A)..(D)` or `(a)..(d)`, also when
split across lines; Statement I/II, Assertion /
Reason, `A.`/`I.`/bullet item lists, "A B C D" True/False grids; Match List-I /
List-II tables (halves cut at a column end are joined, `(I)..(IV)` labels
re-read); a stem with its options misread as a table.

The first run downloads MinerU's models (a few GB). After that a 12-page paper
takes about 2½–3 minutes on an RTX 2050 (≈12 s per page). MinerU's results are cached in `<name>_mineru_basic\`
next to the PDF, so re-running with other layout flags (`--size`, `--columns`,
`--indent`) takes seconds. OCR can misread a formula now and then, so proofread
the answer and solution sections before use.

**Languages.** English papers only for now. The OCR model of the installed
MinerU has no Devanagari characters, so Hindi (or other Indic) text cannot be
read from the page images.

**Regression check.** `tools\score_conversion.py --bench <folder>` re-converts
every `X.pdf` that has its original `X.docx` beside it and scores the result
question by question (found, options, option layout, items, tables, answer key).
The reference set lives in `Paper_Maker_Archive\conversion_bench`.

Test papers, prototypes, checkpoints, and previously generated documents are
kept outside this repository in `Paper_Maker_Archive`.

