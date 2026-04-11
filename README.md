# consolidate-grades

Consolidate student grades from multiple teaching assistant files into a single
Moodle-compatible CSV.

Built for the kind of real-world messiness you get when a dozen TAs each export
grades in their own format: CSV with semicolons, XLSX from LibreOffice with
UTF-7 encoding, ODS from Google Sheets, column names in French or English,
commas instead of dots for decimals, `/20` suffixes, `ABS` markers, typos in
names, multi-tab workbooks with future exams, merged first+last name columns…

The philosophy: **the first time you encounter any quirk, the script asks
you. Your answer is saved to the YAML config so you're never asked again.**
Re-running on a settled config produces zero prompts.

## Features

- **Reads CSV, XLSX, and ODS** — auto-detects CSV encoding (UTF-8, CP1252,
  Latin-1, UTF-7) and separator (comma, semicolon, tab). Strips BOM markers,
  zero-width spaces, and Win-1252 control characters that leak from latin-1
  misreads.
- **Smart column detection** — recognises 30+ column name variants in French
  and English (`Numéro étudiant`, `N°étudiant`, `Student ID`, `NIP`, `Prénom`,
  `First Name`, `Note`, `Grade`, `Résultat`, `Total`…).
- **Multi-sheet workbooks** — when an XLSX/ODS has multiple tabs, the script
  asks which to process; the selection is saved per file. Per-sheet column
  overrides too.
- **Merged name columns** — when a TA puts the full name in a single column
  (`Nom étu`, `Nom Prénom`, `Étudiant`…), the master is indexed in both orders
  and matched anyway.
- **Header row detection** — if a TA puts a title or merged header row above
  the real columns, the script scans the first 10 rows for the actual headers.
- **Robust student matching** — primary match on student ID with edit-distance
  name cross-check; falls back to normalised first+last name; further falls
  back to swapping first/last (TAs sometimes reverse them); finally tries the
  merged-name index.
- **Flexible grade parsing** — handles `14,5` / `14.5` / `15/20`. Built-in
  absent tokens (`ABS`, `ABSENCE`, `DEF`, `DÉFAILLANT`…) treated as no grade,
  and unknown non-numeric values trigger a "does this mean absent?" prompt
  whose answer is persisted.
- **Min/max grade clamping** — optional `min_grade`/`max_grade` in the YAML.
  Out-of-range grades trigger a clamping prompt; the answer persists per file.
- **Safety warnings** — duplicate grades, ID/name mismatches, unknown students,
  ambiguous columns, unparseable values — all reported with colored output.
- **Moodle-ready output** — clean CSV with French column names by default
  (`Numéro d'identification`, `Adresse de courriel`, `Prénom`, `Nom de famille`,
  `<exam_name>`), all customizable. Absent students get empty cells.

## Installation

```bash
git clone https://github.com/fbouchet/grade-consolidator.git
cd grade-consolidator

python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows

pip install .
# Or, for development (editable + test/lint tools):
make dev   # equivalent to: pip install -e ".[dev]"
```

Requires Python 3.10+.

## Quick start

### 1. Create a config file

Minimum required:

```yaml
# config.yaml
master_file: "students_master.csv"

# Option A: list files explicitly
grade_files:
  - "group1_grades.csv"
  - "group2_grades.xlsx"
  - "group3_grades.ods"

# Option B: point to a directory (all .csv/.xlsx/.ods/.tsv are picked up)
grade_dir: "ta_grades/"

# Option C: both — files from grade_files + everything in grade_dir (deduped)

output_file: "moodle_import.csv"
```

Paths are resolved relative to the config file's directory. You must specify
at least one of `grade_files` or `grade_dir`.

On first run, the script will also prompt you for an exam name and a student
ID column name to use as headers in the output CSV. Your answers (and any
other interactive choices like multi-sheet selections, ambiguous columns,
absent tokens, name mismatches, clamping confirmations) are saved back to
the YAML so subsequent runs are fully unattended.

See [`config_example.yaml`](config_example.yaml) for the full annotated
template.

### 2. Run

```bash
consolidate-grades config.yaml
```

Or without installing:

```bash
python -m consolidate_grades.consolidate config.yaml
```

### 3. Check the summary

```
============================================================
CONSOLIDATION SUMMARY
============================================================
  Exam name              : Partiel Mars 2026
  ID column name         : Numéro d'identification
  Master roster          : 120 students
  TA files processed     : 5
  TA files skipped       : 0
  Students with grade    : 115
  Students absent        : 3
  Students without grade : 2
  Output file            : moodle_import.csv

  [OK] group1_grades.csv
    matched=24  grades=23  absent=1
    Using saved override for grade: 'Note /20'.

  [OK] group2_grades.xlsx
    matched=25  grades=25  absent=0
    Using saved sheet selection: ['DC-1', 'DC-2'] (from 3 available).
    Processing sheet 'DC-1'...
    Processing sheet 'DC-2'...

  Students without any grade (2):
    - Martin, Lucas (ID=21045678)
    - Zhang, Wei (ID=21098765)
============================================================
```

## Full config reference

```yaml
# Required
master_file: students_master.csv
grade_files: [notes_A.csv, notes_B.xlsx]   # or grade_dir: ta_grades/
output_file: moodle_import.csv

# Output column headers (prompted on first run)
exam_name: "Partiel Mars 2026"
id_column_name: "Numéro d'identification"

# Grade bounds (optional — triggers clamping prompts when exceeded)
min_grade: 0
max_grade: 20

# Extra absent tokens — auto-populated as you confirm them interactively
absent_tokens: ["-", "NR", "non rendu"]

# --- Below this line: auto-generated as you make interactive choices ---

column_overrides:
  notes_A.csv:
    grade: "Note /20"          # ambiguity resolved interactively
    clamp_above: true          # silently clamp grades > max_grade
  notes_B.xlsx:
    selected_sheets: ["DC-1", "DC-2"]   # which tabs to process
    sheet_columns:                       # per-sheet overrides
      DC-1: {grade: "Note /20"}
      DC-2: {grade: "Note /20"}

name_confirmations:
  notes_A.csv:
    - "21441829"   # ID had a name typo, you confirmed it's the same student
```

## How column detection works

Each column in every file is normalised (lowercased, accents stripped,
hyphens/underscores → spaces, apostrophe-like characters → spaces, C1
control characters dropped, `°`/`º` → `o`) and compared against a set of
known aliases:

| Role       | Example aliases                                          |
|------------|----------------------------------------------------------|
| Student ID | `numéro étudiant`, `n°étudiant`, `student id`, `nip`, `identifiant` |
| First name | `prénom`, `first name`, `given name`                     |
| Last name  | `nom`, `nom de famille`, `last name`, `family name`      |
| Full name  | `nom étu`, `nom prénom`, `étudiant`, `full name`         |
| Email      | `email`, `courriel`, `adresse email`, `e-mail`           |
| Grade      | `note`, `notes`, `grade`, `score`, `résultat`, `note finale`, `total` |

If no grade column matches by name, the tool tries prefix matching
(`Note /20`, `Notes sur 23`), then content-based detection on remaining
columns (≥ 50 % parseable values, excluding `Q1`-style sub-score columns).
When multiple columns match the same role (e.g. both `Numero` and
`Numero etudiant` exist), the tool prompts you to pick one and saves the
choice in the YAML.

## How student matching works

```
1. If the TA file has a student ID column:
   a. Look up the ID in the master roster.
   b. If found AND name columns exist → cross-check the full concatenated
      name. On mismatch, show edit distance and prompt for confirmation
      (saved per ID under name_confirmations so re-runs are silent).
   c. If ID not found → fall back to name matching (step 2).

2. If matching by first+last name (primary or fallback):
   a. Normalise: lowercase, strip accents, hyphens → spaces.
   b. Exactly one match → assign grade.
   c. If no match, try swapping first/last (some TAs reverse them).
   d. Multiple matches (ambiguous) → skip row, warn for manual check.
   e. No match → skip row, warn.

3. If the TA file has a merged "Nom étu"-style column:
   a. Try matching against the master's full-name index, which contains
      every student in both "first last" and "last first" orders.
```

## Development

```bash
make dev       # install in editable mode with dev deps
make test      # run pytest (151 tests)
make lint      # run ruff check
make format    # ruff format + auto-fix
make clean     # remove build artifacts
```

## Design notes

- **Single-module architecture** in `src/consolidate_grades/consolidate.py`
  (~1700 lines). Intentionally not split for teaching/auditing clarity.
- **Backwards-compatible YAML schema** — every new feature adds optional
  keys; old configs keep working unchanged.
- **All interactive choices persist** — re-running the script on a settled
  config produces zero prompts and identical output.
- **Colored terminal output** auto-disables when stdout is not a TTY (CI,
  pipes, redirects).

## License

[MIT](LICENSE)