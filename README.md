# PLM Usage and Licence Optimisation Analysis

Python application for analysing PLM login activity, ADU licence-denial activity, and Production Technician licence usage.
It produces a business-readable Excel workbook focused on licence review and reallocation decisions.

## What It Does

- reads a PLM login audit workbook
- reads an ADU insufficient-licence audit workbook
- optionally reads a Production Technician list
- cleans and validates login data
- groups activity by user
- classifies users as `Regular`, `Occasional`, or `Rare`
- identifies ADU users who may need dedicated licences
- identifies Production Technicians with no or rare observed PLM authentication
- writes a formatted Excel output workbook with executive, recommendation, analysis, and source-data tabs

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
    ├── adu.py
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
- ADU workbook:
  `/Users/frankbogle/Documents/adu/Insufficient ADU License-Login Audit Repot.xlsx`
- ADU sheet:
  `AuditReportExport (1)`
- output workbook:
  `output/plm_usage_and_licence_optimisation_analysis.xlsx`

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
  --adu-input-file "/Users/frankbogle/Documents/adu/Insufficient ADU License-Login Audit Repot.xlsx" \
  --production-technicians-file /Users/frankbogle/Documents/pr_tech/Production_Technician-22-04-26.xlsx \
  --production-technicians-sheet Sheet1 \
  --production-technicians-name-column User \
  --output-file output/plm_usage_and_licence_optimisation_analysis.xlsx
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
- `--adu-input-file`: path to the ADU denial audit workbook
- `--adu-input-sheet`: ADU denial worksheet name
- `--adu-user-column`: ADU audit user column name
- `--adu-timestamp-column`: ADU audit timestamp column name
- `--adu-event-label-column`: ADU audit event label column name
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
`Occasional_Users`, and `Rare_Users`.

Production Technicians not matched to the PLM login audit after name matching are treated as having
no observed PLM authentication during the reporting period.

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

- `Executive_Summary`
- `Licence_Recommendations`
- `Usage_Analysis`
- `ADU_Analysis`
- `Production_Tech_Review`
- `User_Summary`
- `Regular_Users`
- `Occasional_Users`
- `Rare_Users`
- `Reporting`
- `ADU_Denied_Users`
- `Monthly_Usage`
- `Monthly_ADU_Denials`
- `Raw_Login_Data`
- `Raw_ADU_Data`
- `Source_Lists`
- `Rules_And_Assumptions`

## Notes

- `Event Time` values with `BST` and `GMT` suffixes are handled explicitly during cleaning.
- The app derives a cleaner `user_display_name` for workbook readability while preserving the raw user field for traceability.
- Default categorisation uses average distinct active days per month over the full reporting window.
