# FP&A Text Ingestion Starter

This pipeline converts mixed FP&A source files (PowerPoint, Excel, screenshots) into:

- `records.jsonl`: one normalized extracted record per slide/sheet/image
- `chunks.jsonl`: chunked records suitable for embedding + RAG

---

## 1) First-time setup (beginner friendly)

> These steps are safe to copy/paste one line at a time.

### Using Visual Studio / Visual Studio Code

If you are running this from Visual Studio tools:

1. Open the project folder (`Designme`) in your IDE.
2. Open **Terminal** inside the IDE (`View -> Terminal` in VS Code, or Developer PowerShell/Terminal in Visual Studio).
3. Run the same commands below inside that terminal.

On Windows, if `python3` does not work, use `python` or `py` instead.

### A. Check Python is installed

```bash
python3 --version
```

If that fails, try:

```bash
python --version
```

You should see Python 3.10+.

### B. Go to this project

```bash
cd /workspace/Designme
```

### C. Create a virtual environment (recommended)

```bash
python3 -m venv .venv
```

Visual Studio/Windows alternatives:

```powershell
python -m venv .venv
# or
py -m venv .venv
```

### D. Activate it

macOS/Linux:

```bash
source .venv/bin/activate
```

Windows (PowerShell):

```powershell
.venv\Scripts\Activate.ps1
```

### E. Install Python packages

```bash
pip install -r fpna_ingest/requirements.txt
```

If `pip` is not found in Visual Studio terminal, use:

```powershell
python -m pip install -r fpna_ingest/requirements.txt
```

### F. (Optional but recommended) install Tesseract for OCR images

Image OCR needs the **Tesseract** system app in addition to Python packages.

macOS (Homebrew):

```bash
brew install tesseract
```

Ubuntu/Debian:

```bash
sudo apt-get update && sudo apt-get install -y tesseract-ocr
```

Windows:
- Install “Tesseract OCR” (UB Mannheim build is common).
- After install, restart terminal so `tesseract --version` works.

Verify:

```bash
tesseract --version
```

---

## 2) Put your files in one folder

You can put files in **any folder you want**. The folder name is **not required** to follow any special convention.

- ✅ OK: `/Users/you/Documents/fpna_raw`
- ✅ OK: `/workspace/Designme/my-fpna-files`
- ✅ OK: `D:\FPNA\QBR_inputs`

The script reads whatever path you pass as `input_dir` in the run command.

Example:

```text
/my-fpna-files
  /EMEA
    q1_deck.pptx
    forecast.xlsx
    screenshot1.png
```

---

## 3) Run the ingestion

```bash
python fpna_ingest/ingest.py /path/to/my-fpna-files --out-dir ./fpna_out
```

Visual Studio/Windows common form:

```powershell
py fpna_ingest/ingest.py C:\path\to\my-fpna-files --out-dir .\fpna_out
```

If your files are inside this repo, a relative path is fine too:

```bash
python fpna_ingest/ingest.py ./my-fpna-files --out-dir ./fpna_out
```

If your system uses `python3` instead:

```bash
python3 fpna_ingest/ingest.py /path/to/my-fpna-files --out-dir ./fpna_out
```

---

## 4) Check outputs

After the run, you should see:

- `fpna_out/records.jsonl`
- `fpna_out/chunks.jsonl`

Quick preview:

```bash
head -n 3 fpna_out/records.jsonl
head -n 3 fpna_out/chunks.jsonl
```

---

## Supported file types

- **PowerPoint (`.pptx`)**: shape text, table rows, speaker notes
- **Excel (`.xlsx`, `.xlsm`)**: visible cell values per sheet
- **PDF (`.pdf`)**: text extraction per page
- **Images (`.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`)**: OCR text

> `.xls` is not supported in this starter. Convert `.xls` to `.xlsx` first.

---

## Output schema

### `records.jsonl`

Each line includes:

- `doc_id`
- `source_file`
- `source_type` (`ppt_slide`, `excel_sheet`, `pdf_page`, `image_ocr`)
- `location` (for example `slide=4` or `sheet=Revenue`)
- `text`
- `period` (heuristic, for example `2025-Q1`)
- `entity` (heuristic from parent folder name)
- `extracted_at_utc`

### `chunks.jsonl`

Each line includes:

- `chunk_id`
- `doc_id`
- `source_file`
- `source_type`
- `location`
- `chunk_index`
- `text`
- `period`
- `entity`

---

## Troubleshooting

- **`ModuleNotFoundError`**: re-run `pip install -r fpna_ingest/requirements.txt` (inside activated `.venv`).
- **`python: command not found`**: use `python3`.
- **OCR warnings/failures**: install Tesseract and verify with `tesseract --version`.
- **No records extracted**: confirm files are `.pptx`, `.xlsx/.xlsm`, `.pdf`, or supported images.
- **Scanned PDF has no text**: this starter extracts embedded PDF text only; run page images through OCR if needed.
