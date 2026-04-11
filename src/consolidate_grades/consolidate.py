#!/usr/bin/env python3
"""
consolidate_grades.py
=====================
Consolidate student grades from multiple TA files (CSV / XLSX / ODS)
into a single Moodle-compatible CSV, using a master student roster as reference.

Usage:
    consolidate-grades config.yaml
    python -m consolidate_grades.consolidate config.yaml
"""

import argparse
import contextlib
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import yaml

# ============================================================================
# 0. COLORED OUTPUT
# ============================================================================


class _Colors:
    """ANSI color helpers, auto-disabled when stdout is not a terminal."""

    def __init__(self) -> None:
        self.enabled = hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

    def _wrap(self, code: str, text: str) -> str:
        if not self.enabled:
            return text
        return f"{code}{text}\033[0m"

    def info(self, text: str) -> str:
        return self._wrap("\033[36m", text)  # cyan

    def ok(self, text: str) -> str:
        return self._wrap("\033[32m", text)  # green

    def warn(self, text: str) -> str:
        return self._wrap("\033[33m", text)  # yellow

    def error(self, text: str) -> str:
        return self._wrap("\033[31m", text)  # red

    def bold(self, text: str) -> str:
        return self._wrap("\033[1m", text)

    def dim(self, text: str) -> str:
        return self._wrap("\033[2m", text)


C = _Colors()


# ============================================================================
# 1. COLUMN ALIASES & NORMALISATION
# ============================================================================


COLUMN_ALIASES: dict[str, list[str]] = {
    "id": [
        "numero",
        "numero etudiant",
        "numero d etudiant",
        "no etudiant",
        "no d etudiant",
        "n etudiant",
        "num etudiant",
        "code etudiant",
        "numero d identification",
        "no d identification",
        "identifiant",
        "nip",
        "id",
        "student id",
        "studentid",
    ],
    "first_name": [
        "prenom",
        "first name",
        "firstname",
        "given name",
    ],
    "last_name": [
        "nom",
        "nom de famille",
        "last name",
        "lastname",
        "surname",
        "family name",
    ],
    "full_name": [
        "nom etu",
        "nom etudiant",
        "nom et prenom",
        "prenom et nom",
        "nom prenom",
        "prenom nom",
        "nom complet",
        "full name",
        "fullname",
        "etudiant",
        "student",
        "name",
    ],
    "email": [
        "email",
        "e-mail",
        "courriel",
        "adresse de courriel",
        "adresse mail",
        "adresse email",
        "mail",
    ],
    "grade": [
        "note",
        "notes",
        "grade",
        "score",
        "resultat",
        "mark",
        "points",
        "note finale",
        "total",
        "total general",
    ],
}


def normalize_text(text: str) -> str:
    """
    Normalise a string for matching purposes:
    - strip & lowercase
    - strip C1 control characters (U+0080-U+009F) that can leak from
      latin-1 misreads of cp1252 files
    - replace ° and º with 'o' (French N° = Numéro convention)
    - decompose unicode and drop combining characters (accents)
    - replace all dash variants, underscores, and apostrophe-like characters
      with spaces
    - collapse whitespace
    """
    text = text.strip().lower()
    text = re.sub(r"[\u0080-\u009F]", " ", text)  # strip C1 controls
    # ° / º → 'o', inserting a space before the next character if needed
    # (e.g. "N°étudiant" → "no étudiant", not "noétudiant")
    text = re.sub(r"[°º](\S)", r"o \1", text)
    text = re.sub(r"[°º]", "o", text)
    nfkd = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in nfkd if not unicodedata.combining(c))
    # Dashes: hyphen-minus, hyphen, non-breaking hyphen, figure dash,
    #         en dash, em dash, horizontal bar
    # Apostrophes: ASCII quote, grave accent, acute accent, curly quotes,
    #              primes, modifier letters, fullwidth
    text = re.sub(
        r"[-_"
        r"\u2010\u2011\u2012\u2013\u2014\u2015"  # dashes
        r"'\u0060\u00B4"  # basic apostrophes + acute
        r"\u2018\u2019\u201A\u201B"  # curly quotes
        r"\u2032\u2035\u02B9\u02BC"  # primes + modifiers
        r"\uFF07"  # fullwidth
        r"]",
        " ",
        text,
    )
    text = re.sub(r"\s+", " ", text).strip()
    return text


# Characters to strip from the start/end of column names
_COLUMN_GARBAGE = "\ufeff\u200b\u200c\u200d "


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Strip BOM markers, zero-width spaces, and surrounding whitespace."""
    df.columns = [str(col).strip(_COLUMN_GARBAGE).strip() for col in df.columns]
    return df


def _count_alias_matches(values: list[str]) -> int:
    """Count how many values in the list match any known column alias."""
    all_aliases: set[str] = set()
    for aliases in COLUMN_ALIASES.values():
        for a in aliases:
            all_aliases.add(normalize_text(a))
    count = 0
    for v in values:
        if normalize_text(str(v)) in all_aliases:
            count += 1
    return count


def find_header_row(df: pd.DataFrame) -> int | None:
    """
    Check if the current DataFrame headers look like real headers; if not,
    scan the first 10 data rows for one with more alias matches.
    Returns the 0-based data row index to promote, or None if current headers
    are already best.
    """
    current_matches = _count_alias_matches(
        [str(c) for c in df.columns if str(c).strip() and str(c).lower() != "nan"]
    )
    if current_matches >= 2:
        return None

    best_row: int | None = None
    best_count = current_matches

    for i in range(min(10, len(df))):
        row_values = [
            str(v).strip()
            for v in df.iloc[i]
            if pd.notna(v) and str(v).strip() and str(v).lower() != "nan"
        ]
        count = _count_alias_matches(row_values)
        if count > best_count:
            best_count = count
            best_row = i

    return best_row


def promote_header_row(df: pd.DataFrame, row_idx: int) -> pd.DataFrame:
    """Promote a data row to column headers and drop all rows above it."""
    new_headers = [str(v).strip() if pd.notna(v) else "" for v in df.iloc[row_idx]]
    df = df.iloc[row_idx + 1 :].reset_index(drop=True)
    df.columns = new_headers
    return df


def detect_column(columns: list[str], role: str) -> str | None:
    """Return the first column matching the role, or None."""
    aliases = {normalize_text(a) for a in COLUMN_ALIASES.get(role, [])}
    for col in columns:
        if normalize_text(col) in aliases:
            return col
    return None


def detect_all_columns(columns: list[str], role: str) -> list[str]:
    """Return all columns matching the role."""
    aliases = {normalize_text(a) for a in COLUMN_ALIASES.get(role, [])}
    return [col for col in columns if normalize_text(col) in aliases]


# ============================================================================
# 2. GRADE COLUMN DETECTION
# ============================================================================


_QUESTION_PATTERN = re.compile(r"^q\d+$")


def _looks_numeric_grade(series: pd.Series) -> bool:
    """Check whether a column looks like grades (numeric values 0-25 or absent)."""
    non_null = series.dropna()
    if len(non_null) == 0:
        return False
    valid = 0
    for v in non_null:
        s = str(v).strip().replace(",", ".")
        if s == "" or s.lower() == "nan":
            continue
        if s.upper() in ABSENT_TOKENS:
            valid += 1
            continue
        try:
            f = float(s)
            if 0 <= f <= 25:
                valid += 1
        except ValueError:
            pass
    return valid >= max(1, int(0.5 * len(non_null)))


def detect_grade_column(
    df: pd.DataFrame, known_columns: set[str]
) -> tuple[str | None, list[str], list[str]]:
    """
    Detect the grade column.
    Returns (column_name, warnings, ambiguous_candidates).
    """
    warnings: list[str] = []
    cols = [c for c in df.columns if c not in known_columns]

    # 1. Exact alias match
    exact_matches = detect_all_columns(cols, "grade")
    # Prefer "Total" over Q1-Q14 sub-scores
    if exact_matches:
        total_matches = [
            m for m in exact_matches if normalize_text(m) in ("total", "total general")
        ]
        if total_matches:
            return total_matches[0], warnings, []
        if len(exact_matches) == 1:
            return exact_matches[0], warnings, []
        # Multiple exact matches → ambiguous
        return None, warnings, exact_matches

    # 2. Prefix match (e.g. "Note /20", "Notes sur 23")
    grade_aliases = [normalize_text(a) for a in COLUMN_ALIASES["grade"]]
    prefix_matches = []
    for col in cols:
        norm = normalize_text(col)
        for alias in grade_aliases:
            if norm.startswith(alias + " ") or norm == alias:
                if col not in prefix_matches:
                    prefix_matches.append(col)
                break

    if len(prefix_matches) == 1:
        warnings.append(f"  Grade column detected by prefix match: '{prefix_matches[0]}'.")
        return prefix_matches[0], warnings, []
    if len(prefix_matches) > 1:
        warnings.append(
            f"  AMBIGUOUS: multiple grade-like columns by prefix: {prefix_matches}."
        )
        return None, warnings, prefix_matches

    # 3. Content-based fallback: find a numeric column that's not a Q1-Q14
    candidates = []
    for col in cols:
        if _QUESTION_PATTERN.match(normalize_text(col)):
            continue
        if _looks_numeric_grade(df[col]):
            candidates.append(col)

    if len(candidates) == 1:
        warnings.append(
            f"  Grade column auto-detected by content: '{candidates[0]}' "
            "(no column name match)."
        )
        return candidates[0], warnings, []
    if len(candidates) > 1:
        warnings.append(f"  AMBIGUOUS: multiple numeric candidate columns: {candidates}.")
        return None, warnings, candidates

    warnings.append("  No grade column found in this file.")
    return None, warnings, []


# ============================================================================
# 3. NAME NORMALISATION & MATCHING
# ============================================================================


def normalize_name(name: str) -> str:
    """Normalise a student name for comparison."""
    return normalize_text(name)


def make_name_key(first: str, last: str) -> str:
    """Canonical key from first + last name."""
    return f"{normalize_name(last)}|{normalize_name(first)}"


def levenshtein_distance(s1: str, s2: str) -> int:
    """Compute the Levenshtein (edit) distance between two strings."""
    if len(s1) < len(s2):
        return levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev_row = list(range(len(s2) + 1))
    for i, c1 in enumerate(s1):
        curr_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = prev_row[j + 1] + 1
            deletions = curr_row[j] + 1
            substitutions = prev_row[j] + (c1 != c2)
            curr_row.append(min(insertions, deletions, substitutions))
        prev_row = curr_row
    return prev_row[-1]


def name_similarity_hint(dist: int, max_len: int) -> str:
    """Human-readable hint about name similarity."""
    if dist == 0:
        return "identical"
    if max_len == 0:
        return "empty"
    ratio = dist / max_len
    if ratio <= 0.15:
        return "very similar (likely typo)"
    if ratio <= 0.35:
        return "somewhat similar"
    return "very different (possibly wrong student)"


def prompt_name_mismatch(
    sid: str,
    master_name: str,
    ta_name: str,
    filename: str,
) -> bool:
    """
    Interactively confirm an ID match despite a name mismatch.
    Returns True to proceed, False to skip.
    """
    norm_master = normalize_name(master_name)
    norm_ta = normalize_name(ta_name)
    dist = levenshtein_distance(norm_master, norm_ta)
    max_len = max(len(norm_master), len(norm_ta))
    hint = name_similarity_hint(dist, max_len)

    msg = f"Name mismatch for ID '{sid}' in '{filename}':"
    print(f"\n  {C.warn(msg)}")
    print(f"    Master : {master_name}")
    print(f"    TA file: {ta_name}")
    print(f"    Edit distance: {dist} ({hint})")

    while True:
        try:
            raw = input("  Proceed with this match? [y/n]: ").strip().lower()
        except EOFError:
            print("  Please enter y or n.")
            continue
        if raw in ("y", "yes"):
            print(f"  -> {C.ok('Confirmed.')}")
            return True
        if raw in ("n", "no"):
            print(f"  -> {C.warn('Skipping this student.')}")
            return False
        print("  Please enter y or n.")


# ============================================================================
# 4. GRADE PARSING
# ============================================================================

ABSENT_TOKENS = {
    "ABS",
    "ABSENCE",
    "ABSENT",
    "ABJ",
    "ABI",
    "DEF",
    "DEFAILLANT",
    "DÉFAILLANT",
}


@dataclass
class ParsedGrade:
    value: float | None = None
    is_absent: bool = False
    warning: str | None = None


def parse_grade(raw) -> ParsedGrade:
    """Parse a raw grade cell value."""
    if raw is None:
        return ParsedGrade()
    s = str(raw).strip()
    if s == "" or s.lower() == "nan":
        return ParsedGrade(warning="empty cell")
    if s.upper() in ABSENT_TOKENS:
        return ParsedGrade(is_absent=True)
    s = s.replace(",", ".")
    try:
        return ParsedGrade(value=float(s))
    except ValueError:
        return ParsedGrade(warning=f"could not parse grade value '{raw}'")


# ============================================================================
# 5. MASTER ROSTER
# ============================================================================


@dataclass
class Student:
    student_id: str
    first_name: str
    last_name: str
    email: str
    grade: float | None = None
    is_absent: bool = False
    grade_source: str | None = None


@dataclass
class MasterIndex:
    by_id: dict[str, Student] = field(default_factory=dict)
    by_name: dict[str, list[Student]] = field(default_factory=dict)
    # Keyed by normalized full name in both orders ("jean dupont" and
    # "dupont jean" both map to the same student) for merged-column matching.
    by_full_name: dict[str, list[Student]] = field(default_factory=dict)
    all_students: list[Student] = field(default_factory=list)


def _normalize_full_name(text: str) -> str:
    """Normalize a combined-name string for matching."""
    return normalize_text(text)


def build_master_index(df: pd.DataFrame) -> tuple[MasterIndex, list[str]]:
    """Build a MasterIndex from the master DataFrame."""
    warnings: list[str] = []
    cols = list(df.columns)

    id_col = detect_column(cols, "id")
    fn_col = detect_column(cols, "first_name")
    ln_col = detect_column(cols, "last_name")
    em_col = detect_column(cols, "email")

    missing = []
    if id_col is None:
        missing.append("id")
    if fn_col is None:
        missing.append("first_name")
    if ln_col is None:
        missing.append("last_name")
    if missing:
        raise ValueError(
            f"Master file missing required column(s): {missing}. Found columns: {cols}"
        )

    index = MasterIndex()
    for _, row in df.iterrows():
        sid = str(row[id_col]).strip()
        if not sid or sid.lower() == "nan":
            continue
        fn = str(row[fn_col]).strip() if pd.notna(row[fn_col]) else ""
        ln = str(row[ln_col]).strip() if pd.notna(row[ln_col]) else ""
        em = str(row[em_col]).strip() if em_col and pd.notna(row[em_col]) else ""

        if sid in index.by_id:
            warnings.append(f"  Duplicate student ID in master: {sid}. Keeping first.")
            continue

        student = Student(student_id=sid, first_name=fn, last_name=ln, email=em)
        index.by_id[sid] = student
        index.all_students.append(student)
        nk = make_name_key(fn, ln)
        index.by_name.setdefault(nk, []).append(student)
        # Also index by full name in both orders to support merged-name columns
        full_fl = _normalize_full_name(f"{fn} {ln}")
        full_lf = _normalize_full_name(f"{ln} {fn}")
        if full_fl:
            index.by_full_name.setdefault(full_fl, []).append(student)
        if full_lf and full_lf != full_fl:
            index.by_full_name.setdefault(full_lf, []).append(student)

    return index, warnings


# ============================================================================
# 6. FILE READING
# ============================================================================


_CSV_ENCODINGS = ["utf-8", "utf-8-sig", "cp1252", "latin-1", "utf-7"]

# Pattern for UTF-7 escape sequences like +AOk- (accented chars)
_UTF7_PATTERN = re.compile(r"\+[A-Za-z0-9+/]+-")


def _has_utf7_sequences(df: pd.DataFrame) -> bool:
    """Check if column names or first row values contain UTF-7 escape sequences."""
    for col in df.columns:
        if _UTF7_PATTERN.search(str(col)):
            return True
    if len(df) > 0:
        for val in df.iloc[0]:
            if pd.notna(val) and _UTF7_PATTERN.search(str(val)):
                return True
    return False


def _read_single_sheet(path: Path, engine: str, sheet: str | int = 0) -> pd.DataFrame:
    """Read a single sheet from an Excel file with standard cleaning."""
    df = clean_column_names(pd.read_excel(path, engine=engine, sheet_name=sheet, dtype=str))
    header_row = find_header_row(df)
    if header_row is not None:
        df = clean_column_names(promote_header_row(df, header_row))
    return df


def _sheet_looks_like_grades(df: pd.DataFrame) -> bool:
    """Check if a sheet has enough alias matches to be grade data."""
    if df.empty or len(df) == 0:
        return False
    return (
        _count_alias_matches(
            [str(c) for c in df.columns if str(c).strip() and str(c).lower() != "nan"]
        )
        >= 2
    )


def read_file_sheets(
    path: str | Path,
) -> list[tuple[str, pd.DataFrame]]:
    """
    Read a file and return a list of (sheet_name, DataFrame) pairs.

    For CSV files, returns a single pair with sheet_name="".
    For XLSX/ODS files, returns all sheets that look like grade data
    (i.e. have at least 2 column alias matches), or all sheets if there's
    only one.
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix in (".csv", ".tsv", ".txt"):
        return [("", read_file(path))]

    if suffix == ".xlsx":
        engine = "openpyxl"
    elif suffix == ".ods":
        engine = "odf"
    else:
        raise ValueError(f"Unsupported file extension: '{suffix}' for file '{path}'.")

    xls = pd.ExcelFile(path, engine=engine)
    sheet_names = xls.sheet_names

    if len(sheet_names) == 1:
        df = _read_single_sheet(path, engine, sheet_names[0])
        return [(sheet_names[0], df)]

    result: list[tuple[str, pd.DataFrame]] = []
    for name in sheet_names:
        df = _read_single_sheet(path, engine, name)
        if _sheet_looks_like_grades(df):
            result.append((name, df))

    return result


def read_file(path: str | Path) -> pd.DataFrame:
    """
    Read a tabular file (.csv, .xlsx, .ods) into a DataFrame.

    CSV: tries several encodings; auto-detects delimiter via Python csv sniffer.
    XLSX/ODS: reads first sheet only (use read_file_sheets for multi-sheet).
    All formats: column names cleaned of BOM/zero-width chars; if first row
    doesn't look like headers, scans first 10 rows for the real header row.
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix == ".xlsx":
        df = clean_column_names(pd.read_excel(path, engine="openpyxl", dtype=str))
    elif suffix == ".ods":
        df = clean_column_names(pd.read_excel(path, engine="odf", dtype=str))
    elif suffix in (".csv", ".tsv", ".txt"):
        df = None
        for enc in _CSV_ENCODINGS:
            try:
                df = clean_column_names(
                    pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)
                )
                break
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        if df is None:
            raise ValueError(f"Could not read '{path}' with any of the attempted encodings.")
        # Post-read UTF-7 detection
        if _has_utf7_sequences(df):
            with contextlib.suppress(UnicodeDecodeError, pd.errors.ParserError):
                df = clean_column_names(
                    pd.read_csv(
                        path,
                        sep=None,
                        engine="python",
                        encoding="utf-7",
                        dtype=str,
                    )
                )
    else:
        raise ValueError(f"Unsupported file extension: '{suffix}' for file '{path}'.")

    # Check if headers are in a later row
    header_row = find_header_row(df)
    if header_row is not None:
        df = clean_column_names(promote_header_row(df, header_row))

    return df


# ============================================================================
# 7. PROCESS A SINGLE TA FILE
# ============================================================================


@dataclass
class FileReport:
    """Report for a single TA file processing."""

    filename: str
    students_matched: int = 0
    students_absent: int = 0
    grades_assigned: int = 0
    warnings: list[str] = field(default_factory=list)
    skipped: bool = False
    new_file_overrides: dict | None = None
    new_name_confirmations: list[str] = field(default_factory=list)


_ROLE_LABELS = {
    "id": "student ID",
    "first_name": "first name",
    "last_name": "last name",
    "email": "email",
    "grade": "grade",
}


def prompt_column_choice(
    candidates: list[str], filename: str, role: str = "grade"
) -> str | None:
    """
    Interactively ask the user to choose among ambiguous column candidates.
    Returns the chosen column name, or None if the user chooses to skip.
    """
    label = _ROLE_LABELS.get(role, role)
    print(f"\n  {C.warn(f'Multiple {label} columns detected in {filename!r}:')}")
    for i, col in enumerate(candidates, 1):
        print(f"    {i}. {col}")
    print(f"    {len(candidates) + 1}. Skip this file")

    while True:
        try:
            raw = input(f"  Choose {label} column [1-{len(candidates) + 1}]: ").strip()
            choice = int(raw)
        except (ValueError, EOFError):
            print("  Please enter a valid number.")
            continue
        if 1 <= choice <= len(candidates):
            chosen = candidates[choice - 1]
            print(f"  -> {C.ok(f'Using {chosen!r}')}")
            return chosen
        elif choice == len(candidates) + 1:
            print(f"  -> {C.warn('Skipping file.')}")
            return None
        else:
            print(f"  Please enter a number between 1 and {len(candidates) + 1}.")


def prompt_sheet_selection(sheet_names: list[str], filename: str) -> list[str]:
    """
    Interactively ask which sheets to process.
    Returns the selected sheet names (may be empty to skip all).
    """
    print(f"\n  {C.warn(f'Multiple sheets found in {filename!r}:')}")
    for i, name in enumerate(sheet_names, 1):
        print(f"    {i}. {name}")
    print("  Enter sheet numbers to process (comma-separated), or 'all', or 'none':")

    while True:
        try:
            raw = input("  Sheets: ").strip().lower()
        except EOFError:
            print("  Please enter a valid selection.")
            continue

        if raw == "all":
            print(f"  -> {C.ok(f'Using all {len(sheet_names)} sheets.')}")
            return list(sheet_names)
        if raw in ("none", "skip", "0"):
            print(f"  -> {C.warn('Skipping all sheets.')}")
            return []

        try:
            indices = [int(x.strip()) for x in raw.split(",")]
        except ValueError:
            print("  Please enter numbers separated by commas, 'all', or 'none'.")
            continue

        if all(1 <= idx <= len(sheet_names) for idx in indices):
            selected = [sheet_names[idx - 1] for idx in indices]
            print(f"  -> {C.ok(f'Using sheets: {selected}')}")
            return selected

        print(f"  Numbers must be between 1 and {len(sheet_names)}.")


def _resolve_column(
    role: str,
    cols: list[str],
    overrides: dict[str, str],
    filename: str,
    interactive: bool,
    report: FileReport,
) -> tuple[str | None, dict[str, str]]:
    """
    Resolve a column for a role: check overrides → detect → prompt on ambiguity.
    Returns (column_name_or_None, new_overrides_to_save).
    """
    new_overrides: dict[str, str] = {}

    if role in overrides:
        override_col = overrides[role]
        if override_col in cols:
            label = _ROLE_LABELS.get(role, role)
            report.warnings.append(f"  Using saved override for {label}: '{override_col}'.")
            return override_col, new_overrides
        report.warnings.append(
            f"  Saved override '{override_col}' for {role} not found in columns. "
            "Falling back to detection."
        )

    matches = detect_all_columns(cols, role)

    if len(matches) == 1:
        return matches[0], new_overrides
    elif len(matches) == 0:
        return None, new_overrides
    else:
        label = _ROLE_LABELS.get(role, role)
        report.warnings.append(f"  AMBIGUOUS: multiple {label} columns: {matches}.")
        if interactive:
            chosen = prompt_column_choice(matches, filename, role)
            if chosen is not None:
                report.warnings.append(
                    f"  {label.capitalize()} column manually selected: '{chosen}'."
                )
                new_overrides[role] = chosen
            return chosen, new_overrides
        report.warnings.append(f"  Using first match: '{matches[0]}' (non-interactive mode).")
        return matches[0], new_overrides


def process_ta_file(
    path: str | Path,
    master: MasterIndex,
    *,
    interactive: bool = True,
    file_overrides: dict | None = None,
    name_confirmations: list[str] | None = None,
) -> FileReport:
    """Process one TA grade file and update the master index in place.

    file_overrides: per-file override dict from YAML. May be flat
    (single-sheet) or have ``selected_sheets`` and ``sheet_columns`` keys
    (multi-sheet).

    name_confirmations: student IDs for which a name mismatch was previously
    confirmed.
    """
    path = Path(path)
    report = FileReport(filename=str(path))
    overrides = file_overrides or {}
    confirmed_ids = set(name_confirmations or [])

    try:
        all_sheets = read_file_sheets(path)
    except Exception as e:
        report.warnings.append(f"  Could not read file: {e}")
        report.skipped = True
        return report

    if not all_sheets:
        report.warnings.append("  No valid sheets found in file.")
        report.skipped = True
        return report

    # --- Sheet selection for multi-sheet files ---
    is_multi = len(all_sheets) > 1
    selected_sheets_changed = False

    if is_multi:
        all_sheet_names = [name for name, _ in all_sheets]
        saved_selection = overrides.get("selected_sheets")

        if saved_selection is not None:
            selected_names = [n for n in saved_selection if n in all_sheet_names]
            report.warnings.append(
                f"  Using saved sheet selection: {selected_names} "
                f"(from {len(all_sheet_names)} available)."
            )
        elif interactive:
            selected_names = prompt_sheet_selection(all_sheet_names, path.name)
            selected_sheets_changed = True
        else:
            selected_names = all_sheet_names
            report.warnings.append(
                f"  Found {len(all_sheet_names)} sheets, processing all (non-interactive)."
            )

        if not selected_names:
            report.warnings.append("  No sheets selected. Skipping file.")
            report.skipped = True
            if selected_sheets_changed:
                report.new_file_overrides = {"selected_sheets": []}
            return report

        sheets = [(n, df) for n, df in all_sheets if n in selected_names]
    else:
        sheets = all_sheets

    # --- Per-sheet column overrides ---
    saved_sheet_columns: dict[str, dict[str, str]] = overrides.get("sheet_columns", {})
    new_sheet_columns: dict[str, dict[str, str]] = {}
    # Shared overrides accumulate interactive choices across sheets
    shared_overrides: dict[str, str] = {}

    for sheet_name, df in sheets:
        sheet_label = f" [sheet '{sheet_name}']" if sheet_name else ""

        if df.empty:
            report.warnings.append(f"  Sheet{sheet_label} is empty. Skipping.")
            continue

        cols = list(df.columns)

        if is_multi:
            per_sheet = saved_sheet_columns.get(sheet_name, {})
            sheet_col_overrides = {**shared_overrides, **per_sheet}
        else:
            sheet_col_overrides = dict(overrides)
            sheet_col_overrides.pop("selected_sheets", None)
            sheet_col_overrides.pop("sheet_columns", None)

        id_col, id_ov = _resolve_column(
            "id", cols, sheet_col_overrides, path.name, interactive, report
        )
        fn_col, fn_ov = _resolve_column(
            "first_name", cols, sheet_col_overrides, path.name, interactive, report
        )
        ln_col, ln_ov = _resolve_column(
            "last_name", cols, sheet_col_overrides, path.name, interactive, report
        )
        # Merged "Nom étu" / "Nom prénom" / "Étudiant" style columns
        full_col, full_ov = _resolve_column(
            "full_name", cols, sheet_col_overrides, path.name, interactive, report
        )

        known_cols = {c for c in [id_col, fn_col, ln_col, full_col] if c is not None}

        # Grade column
        grade_col: str | None = None
        grade_ov: dict[str, str] = {}

        if "grade" in sheet_col_overrides:
            override_grade = sheet_col_overrides["grade"]
            if override_grade in cols:
                report.warnings.append(
                    f"  Using saved override for grade: '{override_grade}'{sheet_label}."
                )
                grade_col = override_grade
            else:
                report.warnings.append(
                    f"  Saved override '{override_grade}' for grade not found "
                    f"in columns{sheet_label}. Falling back to detection."
                )

        if grade_col is None:
            grade_col, gw, ambiguous = detect_grade_column(df, known_cols)
            report.warnings.extend(gw)

            if grade_col is None and ambiguous and interactive:
                grade_col = prompt_column_choice(ambiguous, path.name, "grade")
                if grade_col is not None:
                    report.warnings.append(f"  Grade column manually selected: '{grade_col}'.")
                    grade_ov["grade"] = grade_col

        if grade_col is None:
            report.warnings.append(f"  No grade column found{sheet_label}. Skipping sheet.")
            continue

        # Collect per-sheet overrides
        this_sheet_ov = {**id_ov, **fn_ov, **ln_ov, **full_ov, **grade_ov}
        if this_sheet_ov:
            existing = new_sheet_columns.get(sheet_name, {})
            existing.update(this_sheet_ov)
            new_sheet_columns[sheet_name] = existing
            shared_overrides.update(this_sheet_ov)

        use_id = id_col is not None
        use_name = fn_col is not None and ln_col is not None
        use_full_name = full_col is not None

        if not use_id and not use_name and not use_full_name:
            report.warnings.append(
                f"  Sheet{sheet_label} has neither a student ID column nor "
                "first/last name columns nor a merged full-name column. "
                "Skipping sheet."
            )
            continue

        if sheet_name:
            report.warnings.append(f"  Processing sheet '{sheet_name}'...")

        # Process each row in this sheet
        for row_idx, row in df.iterrows():
            raw_grade = row.get(grade_col)
            parsed = parse_grade(raw_grade)

            student: Student | None = None
            match_desc = ""

            if use_id:
                sid = str(row[id_col]).strip()
                if sid and sid.lower() != "nan":
                    student = master.by_id.get(sid)
                    match_desc = f"ID={sid}"

                    # Cross-check name if possible
                    if student is not None and use_name:
                        fn_raw = str(row[fn_col]).strip()
                        ln_raw = str(row[ln_col]).strip()
                        if fn_raw.lower() != "nan" and ln_raw.lower() != "nan":
                            expected_fn = normalize_name(student.first_name)
                            expected_ln = normalize_name(student.last_name)
                            got_fn = normalize_name(fn_raw)
                            got_ln = normalize_name(ln_raw)
                            # Compare full concatenated names — handles
                            # different first/last splits
                            master_full_norm = f"{expected_fn} {expected_ln}"
                            ta_full_norm = f"{got_fn} {got_ln}"
                            if master_full_norm != ta_full_norm:
                                master_full = f"{student.first_name} {student.last_name}"
                                ta_full = f"{fn_raw} {ln_raw}"

                                if sid in confirmed_ids:
                                    report.warnings.append(
                                        f"  Row {row_idx}: ID '{sid}' name mismatch "
                                        f"(master='{master_full}', "
                                        f"TA='{ta_full}') — "
                                        "previously confirmed, proceeding."
                                    )
                                elif interactive:
                                    confirmed = prompt_name_mismatch(
                                        sid, master_full, ta_full, path.name
                                    )
                                    if confirmed:
                                        report.new_name_confirmations.append(sid)
                                        report.warnings.append(
                                            f"  Row {row_idx}: ID '{sid}' name "
                                            "mismatch confirmed by user."
                                        )
                                    else:
                                        report.warnings.append(
                                            f"  Row {row_idx}: ID '{sid}' name "
                                            "mismatch rejected by user. "
                                            "Skipping row."
                                        )
                                        student = None
                                        continue
                                else:
                                    report.warnings.append(
                                        f"  Row {row_idx}: ID '{sid}' matches "
                                        f"'{master_full}' in master, "
                                        f"but TA file says '{ta_full}'. "
                                        "Proceeding with ID match "
                                        "(non-interactive)."
                                    )

                    if student is None and sid:
                        if use_name:
                            fn_raw = str(row[fn_col]).strip()
                            ln_raw = str(row[ln_col]).strip()
                            report.warnings.append(
                                f"  Row {row_idx}: ID '{sid}' not found in master. "
                                f"Attempting name fallback ({fn_raw} {ln_raw})."
                            )
                        elif use_full_name:
                            full_raw = str(row[full_col]).strip()
                            report.warnings.append(
                                f"  Row {row_idx}: ID '{sid}' not found in master. "
                                f"Attempting full-name fallback ({full_raw})."
                            )
                        else:
                            report.warnings.append(
                                f"  Row {row_idx}: ID '{sid}' not found in master "
                                "and no name columns to fall back on. Skipping row."
                            )
                            continue

            # Fallback to name matching
            if student is None and use_name:
                fn_raw = str(row[fn_col]).strip()
                ln_raw = str(row[ln_col]).strip()
                if fn_raw.lower() == "nan" or ln_raw.lower() == "nan":
                    report.warnings.append(
                        f"  Row {row_idx}: missing name data. Skipping row."
                    )
                    continue

                nk = make_name_key(fn_raw, ln_raw)
                matches = master.by_name.get(nk, [])
                match_desc = f"name='{fn_raw} {ln_raw}'"

                # If no match in original order, try swapping first/last —
                # some TAs put surname in the "Prénom" column and vice versa
                swap_used = False
                if not matches:
                    swapped_nk = make_name_key(ln_raw, fn_raw)
                    swapped_matches = master.by_name.get(swapped_nk, [])
                    if len(swapped_matches) == 1:
                        matches = swapped_matches
                        swap_used = True

                if len(matches) == 1:
                    student = matches[0]
                    if swap_used:
                        report.warnings.append(
                            f"  Row {row_idx}: name '{fn_raw} {ln_raw}' "
                            f"matched after swapping first/last → "
                            f"{student.first_name} {student.last_name}."
                        )
                elif len(matches) > 1:
                    report.warnings.append(
                        f"  Row {row_idx}: name '{fn_raw} {ln_raw}' matches "
                        f"{len(matches)} students in master. Skipping — "
                        "manual check required."
                    )
                    continue
                else:
                    report.warnings.append(
                        f"  Row {row_idx}: {match_desc} not found in master "
                        "(tried swapped order too). Skipping row."
                    )
                    continue

            # Fallback to full-name matching (merged name column)
            if student is None and use_full_name:
                full_raw = str(row[full_col]).strip()
                if not full_raw or full_raw.lower() == "nan":
                    report.warnings.append(
                        f"  Row {row_idx}: missing full-name data. Skipping row."
                    )
                    continue

                nk = _normalize_full_name(full_raw)
                matches = master.by_full_name.get(nk, [])
                match_desc = f"full_name='{full_raw}'"

                if len(matches) == 1:
                    student = matches[0]
                elif len(matches) > 1:
                    report.warnings.append(
                        f"  Row {row_idx}: full name '{full_raw}' matches "
                        f"{len(matches)} students in master. Skipping — "
                        "manual check required."
                    )
                    continue
                else:
                    report.warnings.append(
                        f"  Row {row_idx}: {match_desc} not found in master. Skipping row."
                    )
                    continue

            if student is None:
                continue

            report.students_matched += 1

            # Check for duplicate grading
            if student.grade is not None or student.is_absent:
                prev_src = student.grade_source or "unknown"
                report.warnings.append(
                    f"  WARNING: student '{student.first_name} {student.last_name}' "
                    f"(ID={student.student_id}) already has a grade from "
                    f"'{prev_src}'. Duplicate found in '{path.name}'. "
                    "Keeping first grade."
                )
                continue

            if parsed.warning and not parsed.is_absent and parsed.value is None:
                report.warnings.append(
                    f"  Row {row_idx}: {match_desc} — {parsed.warning}. No grade assigned."
                )
                continue

            if parsed.is_absent:
                student.is_absent = True
                student.grade_source = str(path.name)
                report.students_absent += 1
            elif parsed.value is not None:
                student.grade = parsed.value
                student.grade_source = str(path.name)
                report.grades_assigned += 1

    # Build new file overrides for persistence
    new_ov: dict = {}
    if is_multi:
        if selected_sheets_changed:
            new_ov["selected_sheets"] = [n for n, _ in sheets]
        if new_sheet_columns:
            new_ov["sheet_columns"] = new_sheet_columns
    else:
        for _sheet_name, ov in new_sheet_columns.items():
            new_ov.update(ov)

    if new_ov:
        report.new_file_overrides = new_ov

    if report.students_matched == 0 and report.grades_assigned == 0:
        report.skipped = True

    return report


# ============================================================================
# 8. OUTPUT
# ============================================================================


def write_moodle_csv(
    master: MasterIndex,
    output_path: str | Path,
    *,
    exam_name: str = "Grade",
    id_column_name: str = "Numéro d'identification",
) -> None:
    """
    Write a Moodle-compatible CSV with French column names by default.
    Both column names are configurable for Moodle import.
    """
    output_path = Path(output_path)
    rows = []
    for s in master.all_students:
        grade_str = ""
        if s.is_absent:
            grade_str = "ABS"
        elif s.grade is not None:
            grade_str = f"{s.grade:g}"
        rows.append(
            {
                id_column_name: s.student_id,
                "Adresse de courriel": s.email,
                "Prénom": s.first_name,
                "Nom de famille": s.last_name,
                exam_name: grade_str,
            }
        )

    df = pd.DataFrame(rows)
    df.to_csv(output_path, index=False, encoding="utf-8")


# ============================================================================
# 9. SUMMARY
# ============================================================================


def print_summary(
    master: MasterIndex,
    reports: list[FileReport],
    output_path: str,
    *,
    exam_name: str = "Grade",
    id_column_name: str = "Numéro d'identification",
) -> None:
    """Print a human-readable summary with colors."""
    total_students = len(master.all_students)
    graded = sum(1 for s in master.all_students if s.grade is not None)
    absent = sum(1 for s in master.all_students if s.is_absent)
    no_grade = total_students - graded - absent

    files_ok = sum(1 for r in reports if not r.skipped)
    files_skipped = sum(1 for r in reports if r.skipped)

    print("\n" + C.bold("=" * 60))
    print(C.bold("CONSOLIDATION SUMMARY"))
    print(C.bold("=" * 60))
    print(f"  {C.info('Exam name')}             : {C.bold(exam_name)}")
    print(f"  {C.info('ID column name')}        : {C.bold(id_column_name)}")
    print(f"  {C.info('Master roster')}         : {total_students} students")
    print(f"  {C.info('TA files processed')}    : {files_ok}")
    if files_skipped:
        print(f"  {C.warn('TA files skipped')}     : {files_skipped}")
    else:
        print(f"  {C.info('TA files skipped')}     : 0")
    print(f"  {C.ok('Students with grade')}   : {graded}")
    if absent:
        print(f"  {C.warn('Students absent')}      : {absent}")
    if no_grade:
        print(f"  {C.error('Students without grade')}: {no_grade}")
    else:
        print(f"  {C.ok('Students without grade')}: 0")
    print(f"  {C.info('Output file')}           : {output_path}")

    for r in reports:
        if r.skipped:
            print(f"\n  [{C.error('SKIPPED')}] {r.filename}")
        else:
            print(f"\n  [{C.ok('OK')}] {r.filename}")
            print(
                f"    matched={r.students_matched}  "
                f"grades={r.grades_assigned}  "
                f"absent={r.students_absent}"
            )
        for w in r.warnings:
            wl = w.lower()
            if any(
                kw in wl
                for kw in (
                    "ambiguous",
                    "warning",
                    "not found",
                    "skipping",
                    "rejected",
                    "could not",
                )
            ):
                print(f"    {C.warn(w)}")
            elif any(
                kw in wl for kw in ("saved override", "previously confirmed", "saved sheet")
            ):
                print(f"    {C.dim(w)}")
            else:
                print(f"    {w}")

    missing = [s for s in master.all_students if s.grade is None and not s.is_absent]
    if missing:
        print(f"\n  {C.error(f'Students without any grade ({len(missing)}):')}")
        for s in missing:
            print(f"    {C.warn('-')} {s.last_name}, {s.first_name} (ID={s.student_id})")

    print(C.bold("=" * 60) + "\n")


# ============================================================================
# 10. CONFIG & MAIN
# ============================================================================


SUPPORTED_EXTENSIONS = {".csv", ".tsv", ".txt", ".xlsx", ".ods"}


def resolve_grade_files(
    config_dir: Path,
    grade_files: list[str] | None = None,
    grade_dir: str | None = None,
) -> list[Path]:
    """Build the list of grade file paths from explicit list and/or directory scan."""
    seen: set[Path] = set()
    result: list[Path] = []

    def _add(p: Path) -> None:
        resolved = p.resolve()
        if resolved not in seen:
            seen.add(resolved)
            result.append(p)

    if grade_files:
        for gf in grade_files:
            gf_path = Path(gf)
            if not gf_path.is_absolute():
                gf_path = config_dir / gf_path
            _add(gf_path)

    if grade_dir:
        dir_path = Path(grade_dir)
        if not dir_path.is_absolute():
            dir_path = config_dir / dir_path
        if not dir_path.is_dir():
            raise ValueError(f"grade_dir '{dir_path}' is not a directory.")
        for child in sorted(dir_path.iterdir()):
            if child.is_file() and child.suffix.lower() in SUPPORTED_EXTENSIONS:
                _add(child)

    return result


def load_config(config_path: str | Path) -> dict:
    """Load and validate the YAML configuration file."""
    config_path = Path(config_path)
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    with open(config_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    if not isinstance(cfg, dict):
        raise ValueError("Config file must be a YAML mapping.")
    if "master_file" not in cfg:
        raise ValueError("Config must specify 'master_file'.")

    has_files = "grade_files" in cfg and isinstance(cfg.get("grade_files"), list)
    has_dir = "grade_dir" in cfg and isinstance(cfg.get("grade_dir"), str)

    if not has_files and not has_dir:
        raise ValueError("Config must specify 'grade_files' (list) and/or 'grade_dir' (path).")

    cfg.setdefault("output_file", "grades_consolidated.csv")
    return cfg


def save_config(config_path: str | Path, cfg: dict) -> None:
    """Write the config dict back to YAML."""
    config_path = Path(config_path)
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(cfg, f, default_flow_style=False, allow_unicode=True, sort_keys=False)


def consolidate(
    config_path: str | Path, *, interactive: bool = True
) -> tuple[MasterIndex, list[FileReport]]:
    """Run the full consolidation pipeline."""
    cfg = load_config(config_path)
    config_changed = False

    config_dir = Path(config_path).parent

    master_path = Path(cfg["master_file"])
    if not master_path.is_absolute():
        master_path = config_dir / master_path

    output_path = Path(cfg["output_file"])
    if not output_path.is_absolute():
        output_path = config_dir / output_path

    # --- Exam name and ID column name (prompt once, save to YAML) ---
    exam_name = cfg.get("exam_name")
    id_column_name = cfg.get("id_column_name")

    if exam_name and id_column_name:
        print(f"  {C.info('Exam name')}      : {C.bold(exam_name)}")
        print(f"  {C.info('ID column name')} : {C.bold(id_column_name)}")
    else:
        if interactive:
            # Explanatory preamble: these are global output settings, not
            # tied to any specific TA file. Without this header, it's
            # confusing why the script is asking for an exam name out of
            # the blue.
            print()
            print(C.bold("=" * 60))
            print(C.bold("First-time setup for this config file"))
            print(C.bold("=" * 60))
            print(
                C.dim(
                    "  These two values become the column headers in the\n"
                    "  output CSV (the file you import into Moodle). They\n"
                    "  will be saved to your YAML config so you won't be\n"
                    "  asked again on subsequent runs.\n"
                )
            )
            if not exam_name:
                try:
                    exam_name = input(
                        C.bold("  Enter exam name for the grade column ")
                        + C.dim("(e.g. 'Partiel Mars 2026')")
                        + C.bold(": ")
                    ).strip()
                except EOFError:
                    exam_name = ""
                if not exam_name:
                    exam_name = "Grade"
                    print(C.dim(f"  -> Using default: '{exam_name}'"))
                cfg["exam_name"] = exam_name
                config_changed = True

            if not id_column_name:
                try:
                    id_column_name = input(
                        C.bold("  Enter student ID column name for output ")
                        + C.dim('(e.g. "Numéro d\'identification")')
                        + C.bold(": ")
                    ).strip()
                except EOFError:
                    id_column_name = ""
                if not id_column_name:
                    id_column_name = "Numéro d'identification"
                    print(C.dim(f"  -> Using default: '{id_column_name}'"))
                cfg["id_column_name"] = id_column_name
                config_changed = True
            print(C.bold("=" * 60))
        else:
            exam_name = exam_name or "Grade"
            id_column_name = id_column_name or "Numéro d'identification"

    print(C.info(f"\nReading master file: {master_path}"))
    master_df = read_file(master_path)
    master, master_warnings = build_master_index(master_df)
    if master_warnings:
        print(C.warn("Master file warnings:"))
        for w in master_warnings:
            print(C.warn(w))

    grade_file_paths = resolve_grade_files(
        config_dir,
        grade_files=cfg.get("grade_files"),
        grade_dir=cfg.get("grade_dir"),
    )

    if not grade_file_paths:
        print(C.warn("WARNING: No grade files found. Output will have no grades."))

    all_overrides: dict[str, dict] = cfg.get("column_overrides", {})
    all_name_confirmations: dict[str, list[str]] = cfg.get("name_confirmations", {})

    reports: list[FileReport] = []
    for gf_path in grade_file_paths:
        print(C.info(f"\nProcessing: {gf_path}"))
        file_ov = all_overrides.get(gf_path.name, {})
        file_name_confs = all_name_confirmations.get(gf_path.name, [])
        report = process_ta_file(
            gf_path,
            master,
            interactive=interactive,
            file_overrides=file_ov,
            name_confirmations=file_name_confs,
        )
        reports.append(report)

        if report.new_file_overrides:
            existing = all_overrides.get(gf_path.name, {})
            for key, val in report.new_file_overrides.items():
                if key == "sheet_columns" and "sheet_columns" in existing:
                    for sn, sv in val.items():
                        existing_sc = existing["sheet_columns"].get(sn, {})
                        existing_sc.update(sv)
                        existing["sheet_columns"][sn] = existing_sc
                else:
                    existing[key] = val
            all_overrides[gf_path.name] = existing
            config_changed = True

        if report.new_name_confirmations:
            existing_confs = all_name_confirmations.get(gf_path.name, [])
            existing_confs.extend(report.new_name_confirmations)
            all_name_confirmations[gf_path.name] = existing_confs
            config_changed = True

    if config_changed:
        if all_overrides:
            cfg["column_overrides"] = all_overrides
        if all_name_confirmations:
            cfg["name_confirmations"] = all_name_confirmations
        save_config(config_path, cfg)
        print(C.ok(f"\n  Overrides saved to {config_path}"))

    write_moodle_csv(
        master,
        output_path,
        exam_name=exam_name,
        id_column_name=id_column_name,
    )
    print_summary(
        master,
        reports,
        str(output_path),
        exam_name=exam_name,
        id_column_name=id_column_name,
    )

    return master, reports


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Consolidate student grades into a Moodle-compatible CSV.",
    )
    parser.add_argument("config", help="Path to the YAML configuration file.")
    args = parser.parse_args()
    consolidate(args.config)


if __name__ == "__main__":
    main()