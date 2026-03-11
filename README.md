---

## pst2pdf

Converts a PST file (Outlook mailbox) to one PDF per email. Built for e-discovery workflows — supports sequential Bates stamping burned directly onto each page and an optional manifest CSV index.

### Requirements

- Python 3.10+
- `libpff` — PST parsing native library
  - Ubuntu/Debian: `sudo apt install libpff-dev`
  - Windows/macOS: installed automatically via the `libpff-python` wheel (`pip install libpff-python`)
- Python packages: `pip install -r pst2pdf/requirements.txt`

### Usage

```bash
cd pst2pdf

# Basic conversion
python pst2pdf.py evidence.pst ./output/

# With Bates stamping burned onto every page + manifest index
python pst2pdf.py evidence.pst ./output/ --bates --prefix SMITH --manifest

# Continue a prior production starting at page 500
python pst2pdf.py evidence2.pst ./output/ --bates --prefix SMITH --start-num 500 --manifest

# Extract attachments as separate files alongside PDFs
python pst2pdf.py evidence.pst ./output/ --attachments extract

# Email bodies only, no attachments, merge everything into one PDF
python pst2pdf.py evidence.pst ./output/ --attachments none --merge all.pdf
```

### GUI

A graphical interface is also available:

```bash
cd pst2pdf
python gui.py
```

The GUI provides:

- **File pickers** for the PST file and output directory
- **Basic options** — prefix, attachment mode (embed / extract / none), manifest CSV toggle, merge toggle
- **Advanced panel** (collapsible) — start number, max attachment size, Bates stamping with pad width, deduplication toggle, non-email items toggle
- **Real-time log** output and progress bar during conversion

All options mirror the CLI flags. The conversion runs in a background thread so the UI stays responsive.

### Options

| Flag | Default | Description |
|---|---|---|
| `--prefix PREFIX` | `MSG` | Filename prefix and Bates label prefix |
| `--start-num N` | `1` | Starting number for filenames and Bates page counter |
| `--bates` | off | Burn Bates numbers onto each page |
| `--bates-pad N` | `6` | Zero-pad width for Bates numbers (6 → `000001`) |
| `--manifest` | off | Write `manifest.csv` index to the output directory |
| `--no-dedup` | off | Disable deduplication — by default emails with the same Message-ID are only output once |
| `--include-non-email` | off | Include calendar events, contacts, and other non-email items (excluded by default) |
| `--attachments MODE` | `embed` | How to handle attachments: `embed` — render as extra pages in the PDF; `extract` — save original files alongside PDFs named `MSG_00001_01_filename.ext`; `none` — ignore attachments entirely |
| `--max-attachment-size MB` | `10` | Skip embedding attachments larger than this size in MB (`embed` mode only). Use `0` to disable the limit. |
| `--merge FILENAME` | off | After converting, merge all PDFs into one file in the output directory (e.g. `--merge all.pdf`) |
| `--verbose` | off | Enable debug logging |

### Output

- `SMITH_00001.pdf`, `SMITH_00002.pdf`, ... — one PDF per email, sequentially named
- Each PDF contains a shaded header block (From, To, CC, Date, Subject, Message-ID, PST folder path, attachments) followed by the email body
- Attachment handling depends on `--attachments` mode:
  - **`embed`** (default) — attachments become extra pages in the PDF:
    - Images (jpg, png, gif, bmp, tiff) — scaled to fit and rendered inline
    - Plain text (txt, csv, log, etc.) — rendered as monospaced text pages
    - PDFs — pages appended directly after the email body
    - Other formats (docx, xlsx, etc.) — a notice page is added with the filename
    - Oversized attachments — skipped with a notice page; use `--max-attachment-size` to adjust the threshold
  - **`extract`** — attachments saved as separate files: `MSG_00001_01_contract.docx`, `MSG_00001_02_photo.jpg`, etc.
  - **`none`** — attachments ignored; PDFs contain headers and body only
- When `--bates` is on, every page (including embedded attachment pages) has a stamped label (e.g. `SMITH_000042`) in the bottom-right corner
- `--merge all.pdf` produces a single PDF of the entire PST alongside the individual files
- `manifest.csv` (optional) — one row per email with columns: `filename`, `date`, `from`, `to`, `cc`, `subject`, `message_id`, `folder_path`, `attachments`, and (if `--bates`) `bates_begin`, `bates_end`

### Build standalone binaries

```bash
# Linux/Mac
./pst2pdf/build.sh        # → pst2pdf/dist/pst2pdf      (CLI)
                          # → pst2pdf/dist/pst2pdf-gui  (GUI)

# Windows
pst2pdf\build.bat         # → pst2pdf\dist\pst2pdf.exe      (CLI)
                          # → pst2pdf\dist\pst2pdf-gui.exe  (GUI)
```

Requires PyInstaller (`pip install pyinstaller`). Both binaries include all dependencies and can be run without a Python installation. The GUI binary uses `--windowed` so no console window appears on macOS/Windows.

Test data shamelessly stolen from: 
https://github.com/aspose-email/Aspose.Email-Python-Dotnet/tree/c564549cfb3b3b1e3d1275dbc1fd01aa5696df35/Examples/Data
