# consolidate-grades

Consolidate student grades from multiple teaching assistant files into a single
Moodle-compatible CSV.

Built for the kind of real-world messiness you get when a dozen TAs each export
grades in their own format: CSV with semicolons, XLSX from LibreOffice, ODS from
Google Sheets, column names in French or English, commas instead of dots for
decimals, `/20` suffixes, `ABS` markers…

## Features

- **Reads CSV, XLSX, and ODS** — auto-detects CSV encoding (UTF-8, Latin-1,
  CP1252) and separator (comma, semicolon, tab).
- **Smart column detection** — recognises 30+ column name variants in French and
  English (`Numéro étudiant`, `Student ID`, `NIP`, `Prénom`, `First Name`,
  `Note`, `Grade`, `Résultat`…).
- **Grade column guessing** — falls back to content-based detection when the
  column isn't named conventionally. Warns on ambiguity instead of guessing wrong.
- **Flexible grade parsing** — handles `14,5` / `14.5` / `15/20` / `ABS` /
  `DEF` / empty cells.
- **Student matching** — primary match on student ID with name cross-check;
  falls back to normalised first+last name when IDs are missing.
- **Safety warnings** — duplicate grades, ID/name mismatches, unknown students,
  ambiguous names, unparseable values — all reported clearly.
- **Moodle-ready output** — clean CSV with `Identifier`, `Email address`,
  `First name`, `Last name`, `Grade`.

## Installation

```bash
# Clone and install
git clone https://github.com/<your-username>/grade-consolidator.git
cd grade-consolidator

# Create a virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows

# Install
pip install .

# Or, for development (editable + test/lint tools):
make dev
# equivalent to: pip install -e ".[dev]"
```

Requires Python 3.10+.

## Quick start

### 1. Create a config file

```yaml
# config.yaml
master_file: "students_master.csv"

grade_files:
  - "group1_grades.csv"
  - "group2_grades.xlsx"
  - "group3_grades.ods"

output_file: "moodle_import.csv"
```

Paths are resolved relative to the config file's directory.

### 2. Run

```bash
consolidate-grades config.yaml
```

Or without installing:

```bash
python -m consolidate_grades.consolidate config.yaml
```

### 3. Check the summary

The tool prints a detailed report:

```
============================================================
CONSOLIDATION SUMMARY
============================================================
  Master roster          : 120 students
  TA files processed     : 5
  TA files skipped       : 0
  Students with grade    : 115
  Students absent (ABS…) : 3
  Students without grade : 2
  Output file            : moodle_import.csv

  [OK] group1_grades.csv
    matched=24  grades=23  absent=1

  [OK] group2_grades.xlsx
    matched=25  grades=25  absent=0
    Grade column auto-detected by content: 'Résultat' (no column name match).

  ...

  Students without any grade (2):
    - Martin, Lucas (ID=21045678)
    - Zhang, Wei (ID=21098765)
============================================================
```

## How column detection works

Each column in every file is normalised (lowercased, accents stripped,
hyphens/underscores → spaces) and compared against a set of known aliases:

| Role       | Example aliases                                          |
|------------|----------------------------------------------------------|
| Student ID | `numéro étudiant`, `student id`, `nip`, `identifiant`    |
| First name | `prénom`, `first name`, `given name`                     |
| Last name  | `nom`, `nom de famille`, `last name`, `family name`      |
| Email      | `email`, `courriel`, `adresse email`, `e-mail`           |
| Grade      | `note`, `grade`, `score`, `résultat`, `note finale`      |

If no grade column matches by name, the tool scans remaining columns for numeric
content (≥ 50 % parseable values) and picks the column if there's exactly one
candidate. Multiple candidates → skip with a warning.

## How student matching works

```
1. If the TA file has a student ID column:
   a. Look up the ID in the master roster.
   b. If found AND name columns exist → cross-check names, warn on mismatch.
   c. If ID not found → fall back to name matching (step 2).

2. If matching by name (primary or fallback):
   a. Normalise: lowercase, strip accents, hyphens → spaces.
   b. Exactly one match → assign grade.
   c. Multiple matches (ambiguous) → skip row, warn for manual check.
   d. No match → skip row, warn.
```

## Development

```bash
make dev       # install in editable mode with dev deps
make test      # run pytest
make lint      # run ruff linter
make format    # auto-format with ruff
make clean     # remove build artifacts
```

## License

[MIT](LICENSE)
