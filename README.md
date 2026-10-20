# Moodle Bulk User Import

A small Flask web app that replaces the error-prone "copy and paste into a spreadsheet" routine of
enrolling many people in Moodle courses. Fill in a form (or paste rows / upload a sheet), and the app
validates everything and generates the files you need:

| File | Content | Purpose |
|---|---|---|
| `.csv` | `username,firstname,lastname,email,password,institution,course1,group1` (UTF-8) | Moodle's *Upload users* tool |
| `.pdf` | Unit, Name, E-mail, Login, Password | Credentials list to hand out |
| `.xlsx` | Input data, the PDF list and the CSV columns | Audit copy |

Files are named `COURSE_ORG_dd.mm.yyyy`; if the name already exists it becomes `... (2)`, so nothing is overwritten.

## Features

- Editable table with **paste from Excel/Sheets** (Ctrl+V) or **upload of an .xlsx** (header row detected automatically)
- Validation before generating: required fields, CPF check digits, e-mail format, duplicated CPF, single organization per batch, course must exist in the list
- Row-level error highlighting in the UI
- Course list managed from the UI (stored in `courses.txt`)
- Output folders configurable per file type, with `{year}` placeholder
- Configurable CSV delimiter and default password via `.env`
- Unit tests for the business rules

## Rules applied

- **username / login** = CPF (Brazilian tax ID), digits only, zero-padded to 11
- **firstname** = first word of the name; **lastname** = the rest
- **institution** = `ORG/UNIT IN UPPERCASE`; **group1** = `UNIT IN UPPERCASE`
- **password** = `DEFAULT_PASSWORD` from `.env`
- Quotes, commas, semicolons and line breaks are stripped from inputs so they can't break the CSV

## Running

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt     # Windows: .venv\Scripts\pip
cp .env.example .env                          # adjust folders / password
.venv/bin/python app.py                       # http://127.0.0.1:8533
```

Tests:

```bash
.venv/bin/python -m pytest
```

## Importing into Moodle

*Site administration → Users → Upload users* → attach the `.csv`, choose the same **CSV delimiter**
configured in `.env` (default `;`) and **Encoding: UTF-8**.

## Running as a service (Linux / systemd)

```ini
[Unit]
Description=Moodle Bulk User Import
After=network-online.target

[Service]
Type=simple
User=moodle-import
WorkingDirectory=/opt/moodle-bulk-user-import
ExecStart=/opt/moodle-bulk-user-import/.venv/bin/python app.py
Restart=on-failure

[Install]
WantedBy=multi-user.target
```

Set `HOST=0.0.0.0` in `.env` to reach it from other machines, and put it behind a reverse proxy with
authentication if it is exposed beyond a trusted network — the app itself has no login.

## Project layout

```
app.py             Flask routes and configuration
moodle_import.py   Validation rules and CSV / PDF / XLSX generation
templates/         Form and course management pages
tests/             pytest suite
courses.txt        Course list
```

## Stack

Python · Flask · openpyxl · ReportLab · vanilla JavaScript
