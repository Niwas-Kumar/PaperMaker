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

The build is written to `dist\PaperMaker.exe`. The local `dist` folder is
intentionally excluded from Git because the current executable is larger than
GitHub's normal per-file limit. Attach release builds to a GitHub Release rather
than committing them to the repository.

## Project files

- `paper_maker.py` - desktop interface
- `paper_core.py` - DOCX parsing and generation
- `pdf_core.py` - PDF parsing and Word generation
- `extract_questions.py` - DOCX command-line utility
- `PaperMaker.spec` - PyInstaller build definition

Test papers, prototypes, checkpoints, and previously generated documents are
kept outside this repository in `Paper_Maker_Archive`.

