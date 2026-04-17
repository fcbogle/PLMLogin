# AGENTS.md

## Project Name
PLM Login Analysis

## Purpose
Build a Python application that reads an Excel file containing PLM login records over a defined period, analyses login behaviour by user, classifies users by usage regularity, and outputs a new Excel workbook with summary tabs and category tabs.

The objective is to help identify:
- users who access PLM regularly
- users who access PLM occasionally
- users who rarely access PLM

The solution should be clear, maintainable, configurable, and suitable for extension over time.

---

## Business Context
The source data contains PLM login records for staff over approximately one year.

Each row represents a login event or access event. The analysis should group activity by user and produce a view of how frequently each user uses PLM.

This is intended to support:
- understanding user adoption of PLM
- identifying usage patterns
- highlighting low-usage users
- supporting possible follow-up actions such as training, engagement, or licence review

The analysis should favour **regularity of usage** rather than simply raw volume of login records.

For example, a user who logs in multiple times on one day but rarely returns should not necessarily be classed as a regular user.

---

## Core Functional Requirements

### Input
The application must:
- read a source Excel workbook
- load a sheet containing PLM login records
- allow configurable input file path
- allow configurable input sheet name
- allow configurable source column names

### Expected Source Columns
At minimum, the application should expect:
- a user identifier column
- a login timestamp or login date/time column

Optional columns may include:
- department
- role
- site
- licence type
- business unit

The code should be written so these optional columns can be incorporated later without major refactoring.

### Data Preparation
The application should:
- standardise column names
- parse login timestamps safely
- derive helper columns such as:
  - login_date
  - year
  - month
  - year_month
- trim whitespace from user identifiers
- optionally normalise case for usernames if required
- remove rows with missing essential fields
- log or report invalid rows

### User-Level Analysis
For each user, calculate:
- total_logins
- distinct_login_days
- active_months
- average_logins_per_month
- average_active_days_per_month
- first_login_date
- last_login_date
- longest_gap_between_login_days

If possible, also support:
- median days between activity
- percentage of months active
- last 90 day activity flag

### Categorisation
Assign each user to one of three categories:
- Regular
- Occasional
- Rare

The first implementation should use configurable thresholds based primarily on:
- average active days per month

Suggested default rules:
- Regular: average_active_days_per_month >= 8
- Occasional: average_active_days_per_month >= 2 and < 8
- Rare: average_active_days_per_month < 2

Thresholds must be stored in configuration, not hard-coded inside analysis logic.

### Output Workbook
Create a new Excel workbook containing the following sheets:

1. Overview
2. Raw_Data
3. User_Summary
4. Regular_Users
5. Occasional_Users
6. Rare_Users
7. Monthly_Activity
8. Category_Rules

### Sheet Requirements

#### Overview
Provide a concise summary including:
- total login records
- total unique users
- date range covered
- count of Regular users
- count of Occasional users
- count of Rare users
- average logins per user
- average active days per user

Also include a small category summary table with:
- category
- user count
- percentage of total users

#### Raw_Data
Write either:
- the cleaned source data, or
- the original data plus derived helper columns

This provides traceability.

#### User_Summary
One row per user with all calculated summary fields and assigned category.

#### Regular_Users
Filtered copy of User_Summary where category == Regular.

#### Occasional_Users
Filtered copy of User_Summary where category == Occasional.

#### Rare_Users
Filtered copy of User_Summary where category == Rare.

#### Monthly_Activity
Matrix by user and month.
Prefer distinct active days per month rather than raw login count.
Suggested structure:
- user
- Jan
- Feb
- Mar
- ...
- Dec
- Total

If the data spans non-calendar months or multiple years, use year-month format instead.

#### Category_Rules
Document the rules used for classification, including threshold values loaded from config.

---

## Technical Requirements

### Language and Libraries
Use Python 3.11+ where possible.

Preferred libraries:
- pandas
- openpyxl
- pathlib
- dataclasses
- typing
- logging

Avoid unnecessary dependencies.

### Project Structure
Use a modular structure such as:

- `main.py`
- `config.py`
- `src/loader.py`
- `src/cleaner.py`
- `src/analyser.py`
- `src/categoriser.py`
- `src/excel_writer.py`
- `src/models.py`
- `src/utils.py`

Tests can later be added under:
- `tests/`

### Coding Standards
- use clear function names
- add docstrings to public functions and classes
- use type hints
- keep functions focused
- avoid monolithic scripts where all logic is in one file
- prefer small reusable components
- do not hard-code file paths in business logic
- keep configuration in one place

### Error Handling
The code should:
- fail clearly when required columns are missing
- handle invalid date values gracefully
- provide helpful messages to the user
- log the number of dropped or invalid rows
- validate output path creation

### Output Formatting
Use `openpyxl` to format the output workbook.

Apply:
- bold header row
- frozen panes
- autofilter
- sensible column widths
- date formatting where appropriate
- conditional formatting or simple fill colours for Usage Category if practical

Optional:
- convert sheets to Excel tables
- apply basic styling consistently

---

## Design Principles

### 1. Prioritise regularity over raw volume
A user’s regular engagement with PLM is more important than the total number of login rows.

Distinct login days and active months should carry more weight than total login count.

### 2. Make categorisation configurable
Thresholds will likely evolve after initial review. The categorisation model should therefore be easy to tune.

### 3. Preserve traceability
The output workbook should make it easy to trace summary results back to the cleaned source data.

### 4. Build for extension
The first version should solve the current use case cleanly, but the structure should support future additions such as:
- department summaries
- role-based summaries
- charts
- inactive user flags
- licence utilisation analysis
- trend analysis across multiple files

---

## Suggested Configuration Model

Store configurable values in `config.py`, for example:

- INPUT_FILE
- INPUT_SHEET
- OUTPUT_FILE
- USER_COLUMN
- TIMESTAMP_COLUMN
- OPTIONAL_COLUMNS
- REGULAR_THRESHOLD
- OCCASIONAL_THRESHOLD
- DATE_FORMATS_TO_TRY

If helpful, use a dataclass such as `AppConfig`.

---

## Suggested Functional Components

### loader.py
Responsibilities:
- read Excel input
- load target sheet
- standardise raw column names minimally

### cleaner.py
Responsibilities:
- validate required columns
- parse timestamps
- derive helper columns
- clean user identifiers
- remove invalid rows

### analyser.py
Responsibilities:
- build per-user metrics
- build monthly activity matrix
- calculate overview stats

### categoriser.py
Responsibilities:
- apply classification rules
- return category labels

### excel_writer.py
Responsibilities:
- write all output sheets
- apply formatting
- create workbook structure

### models.py
Possible dataclasses:
- AppConfig
- CategoryThresholds
- AnalysisSummary

---

## Minimum Viable Product
The MVP must:
1. Read a single Excel file
2. Parse user and login timestamp columns
3. Generate user-level summary metrics
4. Categorise users into Regular / Occasional / Rare
5. Output a formatted workbook with the required tabs

Do not over-engineer beyond this in the first pass.

---

## Nice-to-Have Features
After MVP, consider:
- charts in Overview
- top 10 most active users
- top 10 least active users
- inactive users in last 90 days
- separate department summary tab
- configuration via YAML or JSON
- CLI arguments for input and output paths
- unit tests
- logging to file

---

## Assumptions
Unless the source data proves otherwise, assume:
- each row represents a login-related event
- one user may have multiple rows on the same day
- timestamp values may need parsing
- usernames may contain minor inconsistencies such as leading/trailing spaces
- the source workbook may contain more columns than needed

---

## Non-Goals for Initial Version
Do not prioritise the following in version 1 unless required:
- GUI
- web app
- database storage
- native Excel Pivot Table object generation via Excel automation
- integration with PLM systems
- merging multiple input files automatically

The initial focus is a clean Python-based Excel analysis output.

---

## Guidance for Codex
When implementing:
- start with the MVP
- create clean modular files
- use pandas for data transformation
- use openpyxl for workbook formatting
- include clear inline comments where logic may not be obvious
- keep the solution readable for a business analyst or Python learner
- avoid unnecessary abstraction in the first iteration
- prefer explicitness over cleverness

Implementation order:
1. create project structure
2. define configuration
3. implement loading and validation
4. implement cleaning and helper columns
5. implement user summary metrics
6. implement categorisation
7. implement monthly activity matrix
8. implement Excel writer
9. run end-to-end with a sample workbook
10. refine formatting

---

## Example User Summary Fields
The `User_Summary` sheet should aim to contain columns like:

- user
- total_logins
- distinct_login_days
- active_months
- average_logins_per_month
- average_active_days_per_month
- first_login_date
- last_login_date
- longest_gap_days
- usage_category

Optional future columns:
- department
- site
- licence_type
- last_90_days_active
- percentage_months_active

---

## Example Overview Metrics
The `Overview` sheet should aim to contain:
- total login records
- unique users
- start date
- end date
- regular user count
- occasional user count
- rare user count
- average total logins per user
- average distinct login days per user

---

## Deliverable
A Python project that generates a polished Excel workbook showing PLM user login regularity and classifying users into Regular, Occasional, and Rare usage categories.

The output should be business-readable and suitable for review by management.

