# PDF publication batch checker

`pdf_publication_check.py` checks every matching PDF in a directory against
the page margins stored in the publication's authoritative DOCX template.

Place all PDF papers to be checked in the `input` folder. The program writes
all annotated PDFs and reports to the `output` folder.

```text
pdf_check/
|-- input/                 PDF papers to check
|-- output/                generated results
|-- template.docx          authoritative publication template
`-- pdf_publication_check.py
```

For each input PDF it creates:

- `<name>_checked.pdf` - the original PDF with the correct margins boxed in
  blue and anomalously large whitespace highlighted bright yellow at 70%
  opacity.
- `<name>_report.json` - page-by-page margin and whitespace measurements.
- `batch_summary.csv` - a compact result for the entire batch.

By default, large blank areas must be at least 1.5 inches high. The final page
is excluded because trailing whitespace at the actual end of a paper is
usually legitimate. Detection is raster-based, so text, vector art, and
scanned figures all count as page content.

## Run

From this directory:

```powershell
$python = "C:\Users\johnh\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
& $python .\pdf_publication_check.py
```

The default locations are:

```text
Input papers:  .\input\*.pdf
Results:       .\output\
```

The `output` folder is created automatically if it does not already exist.
Existing source PDFs in `input` are never modified.

To override the standard folders when needed:

```powershell
& $python .\pdf_publication_check.py "C:\path\to\incoming" `
  --output-dir "C:\path\to\reviewed"
```

Useful adjustments:

```text
--pattern "*.pdf"                 input filename pattern
--output-dir ".\output"           results folder
--min-whitespace-height 108       points; 108 = 1.5 inches
--columns 2                       expected one- or two-column layout
--include-final-page              also inspect trailing blank space
--margin-tolerance 6              allowed side-margin variance in points
--pdftoppm "C:\...\pdftoppm.exe"  explicit Poppler executable
```

## Requirements

Python packages: `Pillow`, `pdfplumber`, `pypdf`, and `reportlab`.
Poppler's `pdftoppm` executable must also be installed or supplied with
`--pdftoppm`. The included command uses the bundled Codex runtime, which
already provides these dependencies.
