# Resume Contact Extractor

Pulls **Name, Email, Designation and Contact Number** out of resumes, profiles and contact lists (PDF, Word, text) and puts them into an Excel or CSV sheet. It works even when the document never says "Name:" or "Phone:".

It runs entirely on your own computer, with no AI, no API keys and no uploads. All you need is Python.

```
Resumes folder                         extracted_details.xlsx
├── john_cv.pdf            ──►   Name        | Email            | Designation     | Contact Number
├── priya-resume.docx            John Smith  | john@mail.com    | Data Analyst    | +91 98200 12345
└── vendor_contacts.docx         Priya Rao   | priya@mail.com   | HR Executive    | 98450 11223
                                 ...
```

## Features

- **No labels needed:** it works out which line is the name and which is the job title from the layout and other clues.
- **PDF, Word (.docx), and text files.** Old `.doc`, `.rtf` and `.odt` files work too if MS Word or LibreOffice is installed.
- **One person per file** (resumes) or **many people per file** (contact lists or tables, with or without headers), detected automatically.
- **Adds rows to your existing Excel sheet** under its own headers ("Candidate Name", "Mobile No", "Email ID", "Position", "Sr No", …).
- **Skips duplicates** (same email already in the sheet).
- **Flags anything uncertain:** those rows are highlighted yellow, and a log file lists every file that needs a manual check.
- **Installs its own helper packages** on first run.

## Requirements

- Python 3.8 or newer ([download](https://www.python.org/downloads/); on Windows tick **"Add python.exe to PATH"** during setup).
- Internet the **first time only**, so it can install `pymupdf` and `openpyxl`.
  If that fails, run once: `python -m pip install -r requirements.txt`

## Quick start (Windows)

1. Download this repo: **Code → Download ZIP**, then unzip it.
2. Put all the PDF/Word files in one folder.
3. Double-click **`run_extractor.bat`** and choose that folder (or drag the folder onto it).
4. Open **`extracted_details.xlsx`**, which is created inside that folder.

## Command line (any OS)

```bash
python extract_contacts.py                          # a window asks you to pick a folder
python extract_contacts.py "C:\Resumes"             # every file in the folder and its sub-folders
python extract_contacts.py a.pdf b.docx -o out.csv  # specific files, CSV output
python extract_contacts.py "C:\Resumes" -o "C:\MySheet.xlsx"   # add rows to an existing sheet
```

| Option | What it does |
|---|---|
| `-o FILE` | Output `.xlsx` or `.csv`. If it exists, rows are appended under its headers. Default: `extracted_details.xlsx` in the input folder. |
| `--mode single` | Every file is one person (resumes). |
| `--mode multi` | Every file has many people (contact lists, tables). |
| `--no-subfolders` | Don't look inside sub-folders. |
| `--allow-duplicates` | Also add people whose email is already in the sheet. |
| `--show-text` | Print the text read from each file (useful for troubleshooting). |

Close the sheet in Excel before running; the script can't write to a file Excel has open.

## Checking the results

Each run produces two things:

- **The sheet:** rows with a note in the **Needs Review** column are highlighted yellow (e.g. `no phone`, `check name`, `name taken from LinkedIn URL`).
- **`extracted_details_log.txt`**, saved next to the sheet, which lists:
  - **ERRORS**: files that couldn't be opened (damaged, password-protected, …)
  - **NEEDS REVIEW**: what was missing or uncertain in each file, and what was found
  - **DUPLICATES**: people skipped because their email was already there
  - **ALL FILES**: every file marked `OK`, `REVIEW`, `ERROR` or `SKIPPED`

> The extractor uses rules and scoring, so it can't be 100% right. A row marked `OK` was found confidently, but it is still a best guess, so spot-check a few.

## How it works

| Field | How it is found |
|---|---|
| Email | Pattern match. It also repairs addresses split over two lines. |
| Contact number | Pattern match on 10–13 digit numbers. Years ("2014 – 2018"), dates, PIN codes and ID numbers (Aadhaar, passport, …) are ignored, and numbers next to words like *Phone / Mobile / 📞* come first. |
| Name | Each line is scored: does it match the email address, LinkedIn URL or file name? Is it at the top, or close to the phone and email? Does it look like a name, rather than a heading, city or skill? It also handles names split over two lines and letter-spaced names (`J O H N  S M I T H`). |
| Designation | The job-title line just under the name, otherwise the current job in the experience section ("… – Present"). Company names, dates and places are removed from it. |

## Optional extras

- **Scanned PDFs / photos:** install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki), then run `python -m pip install pytesseract pillow`.
- **Different column names for new sheets:** edit `OUTPUT_COLUMNS` at the top of `extract_contacts.py`.

## Privacy

Everything happens locally; no file content is sent anywhere. The `.gitignore` blocks PDF, Word, Excel and CSV files so resumes and results can't be committed by accident.

## License

[MIT](LICENSE)
