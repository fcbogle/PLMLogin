# PLM Login Analysis

Python application for analysing PLM login activity from an Excel export and producing a business-readable Excel workbook.

## What It Does

- reads a source Excel workbook
- cleans and validates login data
- groups activity by user
- classifies users as `Regular`, `Occasional`, or `Rare`
- writes a formatted Excel output workbook with summary tabs

The analysis is designed to prioritise regularity of usage over raw login volume.

## Project Structure

```text
plm_login/
├── main.py
├── config.py
├── config/
│   └── exclusions.json
└── src/
    ├── analyser.py
    ├── categoriser.py
    ├── cleaner.py
    ├── excel_writer.py
    ├── loader.py
    └── models.py
```

## Requirements

- Python 3.11+
- `pandas`
- `openpyxl`

Install dependencies with:

```bash
pip install pandas openpyxl
```

## Default Input and Output

By default the app uses:

- input workbook:
  `/Users/frankbogle/Documents/login/BlatchfordUserLoginAuditReportExport.xlsx`
- input sheet:
  `AuditReportExport`
- output workbook:
  `output/plm_login_analysis.xlsx`

These can be overridden from the command line.

## Run The Application

Use defaults:

```bash
python main.py
```

Specify input and output paths:

```bash
python main.py \
  --input-file /Users/frankbogle/Documents/login/BlatchfordUserLoginAuditReportExport.xlsx \
  --output-file output/plm_login_analysis.xlsx
```

See all options:

```bash
python main.py --help
```

## CLI Options

- `--input-file`: path to the source Excel workbook
- `--input-sheet`: source worksheet name
- `--output-file`: output workbook path
- `--user-column`: source user column name
- `--timestamp-column`: source timestamp column name
- `--exclusions-file`: path to a JSON exclusions file
- `--normalise-user-case`: lower-case user identifiers during cleaning
- `--disable-default-exclusions`: include built-in test/admin accounts
- `--production-technicians-file`: optional Excel or CSV file containing Production Technician full names
- `--production-technicians-sheet`: worksheet name for the Production Technician file
- `--production-technicians-name-column`: column containing Production Technician full names, defaults to `Full Name`

## Production Technician Review

The app can optionally load a simple Production Technician extract and match it against the derived
`user_display_name` from the login audit.

Minimum extract:

```text
Full Name
Andy Self
```

Example run:

```bash
python main.py \
  --production-technicians-file /path/to/production_technicians.xlsx \
  --production-technicians-sheet Sheet1
```

Matched Production Technicians are highlighted in `User_Summary`, `Regular_Users`,
`Occasional_Users`, and `Rare_Users`. The workbook also includes a `Production_Techs`
sheet and a Production Technician match section on `Reporting`.

## Exclusion List

User exclusions are loaded from:

`config/exclusions.json`

Supported JSON fields:

```json
{
  "excluded_user_exact_values": [
    "WPS Test (WPS Test: Blatchford)"
  ],
  "excluded_user_contains_values": [
    "test"
  ]
}
```

- `excluded_user_exact_values`: exclude exact user string matches
- `excluded_user_contains_values`: exclude users where the raw user value contains the given text

If the JSON file is omitted, the app still works.

## Output Workbook Tabs

The generated workbook contains:

- `Overview`
- `Raw_Data`
- `User_Summary`
- `Regular_Users`
- `Occasional_Users`
- `Rare_Users`
- `Monthly_Activity`
- `Production_Techs`
- `Category_Rules`
- `Reporting`

## Notes

- `Event Time` values with `BST` and `GMT` suffixes are handled explicitly during cleaning.
- The app derives a cleaner `user_display_name` for workbook readability while preserving the raw user field for traceability.
- Default categorisation uses average distinct active days per month over the full reporting window.
