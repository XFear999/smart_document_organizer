# Smart Document Organizer

A safe, accurate, **rule-based** document organizer for Windows (also runs on
macOS/Linux). Point it at your Downloads folder and it will read each file,
figure out what it is, detect exact duplicates, and — only when you tell it to —
copy or move everything into a clean, organized folder tree. It never deletes
anything and never overwrites.

---

## Why it's safe

| Guarantee | How it works |
|-----------|--------------|
| **Dry-run by default** | Nothing is touched unless you pass `--apply`. |
| **Copy is the default operation** | Originals are left alone. `--move` must be explicitly requested. |
| **No overwrites** | Name collisions get ` (2)`, ` (3)`, … appended. |
| **No deletion** | The app never deletes a file, ever. |
| **No forced guesses** | Weak/ambiguous matches go to `_Needs_Review` (or are just logged), never jammed into the wrong folder. |
| **Skips system + output folders** | Won't descend into Windows/AppData/Recycle Bin and won't re-organize its own output. |

---

## Install

1. **Install Python 3.9+** from [python.org](https://www.python.org/downloads/)
   (check "Add Python to PATH" during install). Tkinter is included.

2. **Install the Python dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

3. **(Optional) For OCR** of scanned PDFs and images, install two pieces of
   system software and make sure they're on your `PATH`:
   - **Tesseract OCR** — https://github.com/UB-Mannheim/tesseract/wiki
   - **Poppler** (needed to rasterize scanned PDFs) —
     https://github.com/oschwartz10612/poppler-windows

   Without these, the app still runs — it just can't read text out of scanned
   images/PDFs and will route those files to `_Needs_Review`.

---

## Quick start

```powershell
# 1) SAFE PREVIEW — reads everything, writes a report, changes nothing:
python smart_document_organizer.py "C:\Users\me\Downloads"

# 2) Review the report:
#    C:\Users\me\Downloads\Organized_Documents\organizer_log.csv

# 3) When happy, actually COPY files into the organized tree (originals kept):
python smart_document_organizer.py "C:\Users\me\Downloads" --apply --copy --rename
```

Prefer clicking? Launch the GUI:

```powershell
python smart_document_organizer.py --gui
```

---

## Example commands

```powershell
# Preview with OCR for images and scanned PDFs, and standardized renaming:
python smart_document_organizer.py "C:\Users\me\Downloads" --use-ocr --scanned-pdf-ocr --rename

# Copy into a custom output folder, sending low-confidence files to _Needs_Review:
python smart_document_organizer.py "C:\Users\me\Downloads" --apply --copy --move-needs-review --output "D:\Organized"

# Put duplicate copies into _Duplicates (default), or leave them in place:
python smart_document_organizer.py "C:\Users\me\Downloads" --apply --duplicate-action folder
python smart_document_organizer.py "C:\Users\me\Downloads" --apply --duplicate-action skip

# MOVE instead of copy (relocates originals — asks/​warns first):
python smart_document_organizer.py "C:\Users\me\Downloads" --apply --move

# Be stricter about what gets filed vs. sent to review:
python smart_document_organizer.py "C:\Users\me\Downloads" --min-confidence 0.6

# Teach it from your corrections (see below), then apply:
python smart_document_organizer.py "C:\Users\me\Downloads" --learn-from "C:\Users\me\Downloads\Organized_Documents\organizer_log.csv" --apply --copy
```

### All command-line options

| Option | Meaning |
|--------|---------|
| `folder` | Source folder to scan (recursively). |
| `--output`, `-o` | Output folder. Default `<folder>/Organized_Documents`. |
| `--apply` | Actually copy/move. Omit for a dry-run. |
| `--copy` | Copy files (safe, the default operation). |
| `--move` | Move files instead (relocates originals). |
| `--use-ocr` | OCR image files (JPG/PNG/TIFF/WEBP). |
| `--scanned-pdf-ocr` | OCR PDFs that have no extractable text. |
| `--duplicate-action keep\|folder\|skip` | Duplicate policy (default `folder`). |
| `--min-confidence` | 0–1 threshold to file vs. review (default `0.45`). |
| `--move-needs-review` | Also copy/move low-confidence files into `_Needs_Review`. |
| `--learn-from CSV` | Apply `corrected_category` from a CSV by SHA-256. |
| `--rename` | Rename to `YYYY-MM-DD - Category - Party - Hint.ext`. |
| `--ocr-lang` | Tesseract language(s), e.g. `eng` or `eng+ara`. |
| `--gui` | Launch the Tkinter GUI. |
| `--verbose`, `-v` | Verbose logging. |

---

## How classification works (accuracy)

This is **not** naive keyword matching. Each category has a rule-set with four
kinds of signals plus a must-have gate:

- **Must-have signals** — a category is *disqualified* (score forced to 0)
  unless a minimum number of these are present. This is what stops a random
  file from being forced into "Bank Statements" just because it says "balance".
  - *Bank statements* require real banking signals (statement period,
    beginning/ending balance, deposits/withdrawals, account/routing number…).
  - *Invoices* require an invoice number **and** an amount-due/total-due.
  - *Legal filings* require a court context **and** docket/plaintiff/motion/…
  - *Contracts* require agreement/contract **and** whereas/effective date/
    signature/…
  - *IRS notices* require "Internal Revenue Service"/"Department of the
    Treasury", with strong bonus for `Notice CPxx`, `Form 4564`, IDR, audit…
  - *Tax forms* detect `1040/1120/1065/990/W-2/W-9/1099/K-1`.
- **Strong keywords** — high-signal phrases, +2.0 each.
- **Weak keywords** — supporting evidence, +0.6 each.
- **Negative keywords** — evidence *against* the category, −1.5 each.

The best-scoring category wins, and a **confidence** (0–1) is derived from the
raw score. A file is sent to **Needs Review** when confidence is below
`--min-confidence`, when there are no must-have signals, or when the top two
categories are too close to call (ambiguous). Every decision is logged with a
readable score breakdown in the `reason` column.

---

## Folder structure produced

```
Organized_Documents/
  Financial/
    Bank Statements/ <Bank>/ <Year>/
    Credit Card Statements/ <Issuer>/ <Year>/
    Invoices/ <Party>/
    Receipts/ <Party>/
  Tax/
    IRS Notices/
    Tax Forms/
  Legal/
    Contracts/
    Drafts/
    Court Filings/
  Business/
    Business Documents/
    Payroll and HR/
  Personal/
    Identity Documents/
    Immigration and Travel/
  Real Estate/
  Insurance/
  Medical/
  Vehicles/
  Utilities and Bills/ <Party>/
  _Needs_Review/
  _Duplicates/
  organizer_log.csv
  organizer.db
```

---

## How duplicate handling works

1. Every file's **SHA-256** content hash is computed while scanning.
2. The **first** file seen with a given hash is the **original**. Any later
   file with the same hash is an **exact duplicate** (identical bytes), and the
   original is never disturbed.
3. Duplicates are handled per `--duplicate-action`:
   - **`keep`** — organize the duplicate normally alongside the original, but
     never overwrite (collisions get ` (2)`, ` (3)`…).
   - **`folder`** *(default)* — copy/move duplicate copies into `_Duplicates`.
   - **`skip`** — leave duplicates exactly where they are and only log them.
4. Nothing is ever overwritten, and the "Exact Duplicates" relationship
   (`duplicate_status`, `duplicate_of`, `duplicate_group`) is recorded in the
   log so you can audit it.

The `duplicate_group` column is the first 12 hex chars of the SHA-256, so every
copy of the same content shares a group id.

---

## How correction learning works

The organizer is honest about uncertainty, and you can teach it:

1. Run a scan. Open `Organized_Documents/organizer_log.csv`.
2. Find rows you disagree with and type the right category into the
   **`corrected_category`** column (e.g. change a mis-filed row to `Invoices`).
   Leave rows you agree with blank.
3. Save the CSV, then run again with:
   ```powershell
   python smart_document_organizer.py "C:\Users\me\Downloads" ^
       --learn-from "C:\Users\me\Downloads\Organized_Documents\organizer_log.csv" ^
       --apply --copy
   ```
4. Corrections are keyed by **SHA-256**, so the exact same file content is
   classified using your correction — with confidence `1.0` — no matter where
   it now lives or what it's named. Learned matches always win over the rules.

Because it's content-hash based, it's robust to renaming and re-downloading the
same document.

---

## Tracking / audit trail

Both the CSV and the SQLite database (`organizer.db`, table `files`) contain:

`original_path, original_name, original_size, new_path, new_name, category,
confidence, needs_review, detected_date, detected_party, detected_bank,
sha256, duplicate_group, duplicate_status, duplicate_of, reason (score
details), text_preview, corrected_category, status`.

Query the DB directly if you like:
```powershell
python -c "import sqlite3;c=sqlite3.connect(r'C:\Users\me\Downloads\Organized_Documents\organizer.db');[print(r) for r in c.execute('select category,count(*) from files group by category')]"
```

---

## Supported file types

PDF (incl. scanned via OCR), DOCX, DOC *(best-effort)*, XLSX, XLS, CSV, TXT,
JPG, PNG, TIFF, WEBP, EML, MSG, ZIP *(classified by its contents)*. Anything
else is logged and routed to `_Needs_Review` — never moved blindly.

---

## Recommended first workflow

1. `python smart_document_organizer.py "C:\Users\me\Downloads" --use-ocr --scanned-pdf-ocr --rename`  ← preview
2. Skim `organizer_log.csv`; fix any `corrected_category` cells.
3. `... --learn-from "...\organizer_log.csv" --apply --copy --rename`  ← safe copy
4. Verify the `Organized_Documents` tree looks right.
5. Only then, if you want your Downloads folder emptied, re-run with `--move`.
