# CLAUDE.md

## Project overview

`consolidate-grades` is a CLI tool that merges student grades from multiple teaching assistant files (CSV, XLSX, ODS) into a single Moodle-compatible CSV. It is used in a French university context (UL1IN002, first-year C programming course).

## Project structure

```
src/consolidate_grades/
  __init__.py          # Public API re-exports
  consolidate.py       # All logic: column detection, file I/O, matching, output
tests/
  test_consolidate.py  # 104 tests covering all key functions
config_example.yaml    # Sample YAML config for users
```

Single-module architecture — everything is in `consolidate.py`. This is intentional for a small teaching utility; don't split into multiple modules unless it grows significantly.

## Commands

```bash
make dev          # Install in editable mode with dev deps
make test         # Run pytest (equivalent to: pytest -v)
make lint         # Run ruff check
make format       # Run ruff format + ruff check --fix
```

## Key design decisions

- **Column detection** is alias-based (30+ French/English variants) with content-based fallback for grades. Aliases are normalised at comparison time (lowercase, strip accents, hyphens→spaces). When adding new aliases, add them to `COLUMN_ALIASES` in `consolidate.py` and add a parametrized test case in `TestDetectColumn`.
- **Student matching** is by ID first (with name cross-check), fallback to normalised name. Ambiguous names are never silently resolved — they produce warnings.
- **Name normalisation** strips accents and hyphens only. Ligatures (œ, æ) are deliberately preserved — this matches the user's specification.
- **Grade parsing** is permissive: comma/dot decimals, `/20` suffixes, absent tokens (ABS/DEF/ABJ/ABI). No assumption about max grade.
- **All warnings** go to the console summary, never silently swallowed.
- **Grade file resolution** supports explicit `grade_files` list, a `grade_dir` directory scan, or both (deduplicated by resolved path). Directory scan picks up `.csv`, `.xlsx`, `.ods`, `.tsv`, `.txt` and sorts alphabetically for deterministic order.

## CI

GitHub Actions runs on push/PR to main. Matrix: Python 3.10–3.13. Steps: install, ruff check, ruff format --check, pytest. Workflow file: `.github/workflows/ci.yml`.

## Testing conventions

- Tests use `tmp_path` (pytest fixture) for all file I/O — no test touches the real filesystem.
- Helper functions `_master_df()`, `_build_master()`, `_write_csv()` at the top of the test file create standard fixtures.
- When adding a new feature, add both a positive test and at least one edge-case/warning test.

## Linting

Ruff is configured in `pyproject.toml`. Target is Python 3.10+. Key rules: E, W, F, I, N, UP, B, SIM, RUF. Line length 88. Run `make lint` before committing.

## Common tasks

- **Add a new column alias**: update `COLUMN_ALIASES` dict → add parametrized test in `TestDetectColumn` → run `make test`.
- **Add a new absent token**: update `ABSENT_TOKENS` set → add test in `TestParseGrade` → run `make test`.
- **Support a new file format**: add a branch in `read_file()` → add a test in `TestReadFile` → add an integration test.