#!/usr/bin/env python3
"""
consolidate_grades.py
=====================
Consolidate student grades from multiple TA files (CSV / XLSX / ODS)
into a single Moodle-compatible CSV, using a master student roster as reference.

Usage:
    python consolidate_grades.py config.yaml
    python consolidate_grades.py --help
"""

import argparse
import contextlib
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import yaml

# ============================================================================
# 1. COLUMN NAME DETECTION
# ============================================================================

# Canonical aliases for each semantic role.
# All entries are stored *already normalised* (lowercase, no accents, etc.)
COLUMN_ALIASES: dict[str, list[str]] = {
    "id": [
        "id",
        "numero etudiant",
        "num etudiant",
        "num_etudiant",
        "student id",
        "student_id",
        "studentid",
        "n etudiant",
        "no etudiant",
        "no_etudiant",
        "identifiant",
        "nip",
        "code etudiant",
        "code_etudiant",
        "numero",
        "numero d identification",
        "no d identification",
        "numero d etudiant",
        "no d etudiant",
    ],
    "first_name": [
        "prenom",
        "first name",
        "first_name",
        "firstname",
        "given name",
        "givenname",
    ],
    "last_name": [
        "nom",
        "nom de famille",
        "last name",
        "last_name",
        "lastname",
        "family name",
        "family_name",
        "familyname",
        "nom famille",
        "nom_famille",
    ],
    "email": [
        "email",
        "mail",
        "courriel",
        "adresse email",
        "adresse mail",
        "adresse_email",
        "adresse de courriel",
        "e-mail",
        "email address",
    ],
    "grade": [
        "note",
        "notes",
        "grade",
        "score",
        "resultat",
        "result",
        "mark",
        "points",
        "notation",
        "note finale",
        "note_finale",
        "final grade",
        "final_grade",
        "total",
        "total general",
        "note totale",
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
    - collapse whitespace (this also merges "--", "- -", etc. into one space)
    """
    text = text.strip().lower()
    text = re.sub(r"[\u0080-\u009F]", " ", text)  # strip C1 controls
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


# Characters to strip from the start of column names (BOM, zero-width spaces, etc.)
_COLUMN_GARBAGE = "\ufeff\u200b\u200c\u200d\ufffe"


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Strip BOM markers, zero-width characters, and surrounding whitespace
    from all column names.  Returns the DataFrame (modified in place).
    """
    df.columns = [col.strip(_COLUMN_GARBAGE).strip() for col in df.columns]
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
    Check whether the current DataFrame column headers look like real headers.
    If not, scan the first 10 data rows for a row with more alias matches.

    Returns the 0-based *data row index* to promote to header, or None if the
    current headers are already the best candidate.

    This handles files where TAs have put a title, corrector info, or other
    metadata in the rows above the actual column headers.
    """
    current_matches = _count_alias_matches(
        [str(c) for c in df.columns if str(c).strip() and str(c).lower() != "nan"]
    )
    if current_matches >= 2:
        return None  # Current headers look fine

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
    """
    Promote a data row to column headers and drop all rows above it.
    """
    new_headers = [str(v).strip() if pd.notna(v) else "" for v in df.iloc[row_idx]]
    df = df.iloc[row_idx + 1 :].reset_index(drop=True)
    df.columns = new_headers
    return df


def detect_column(columns: list[str], role: str) -> str | None:
    """
    Given a list of DataFrame column names, return the *first* column name
    that matches the semantic ``role``.  Returns ``None`` if no match found.

    Use ``detect_all_columns`` when you need to detect ambiguity.
    """
    aliases = {normalize_text(a) for a in COLUMN_ALIASES.get(role, [])}
    for col in columns:
        if normalize_text(col) in aliases:
            return col
    return None


def detect_all_columns(columns: list[str], role: str) -> list[str]:
    """
    Return *all* column names that match the semantic ``role``.
    Used to detect ambiguity (e.g. both 'Numero' and 'Numero etudiant').
    """
    aliases = {normalize_text(a) for a in COLUMN_ALIASES.get(role, [])}
    return [col for col in columns if normalize_text(col) in aliases]


def detect_grade_column(
    df: pd.DataFrame, known_columns: set[str]
) -> tuple[str | None, list[str], list[str]]:
    """
    Detect the grade column in *df*.

    Strategy:
      1. Exact name-based matching using COLUMN_ALIASES["grade"].
      2. Prefix matching: column name starts with a grade alias followed by
         a separator (e.g. "Note /20", "Score final").
      3. Content-based: find columns that are >= 50 % numeric-like (including
         ABS / DEF / ABJ tokens), excluding columns already assigned a role.
         If multiple candidates, prefer one named "Total" (or similar)
         when sub-score columns (Q1, Q2, ...) are present.

    Returns (column_name_or_None, list_of_warnings, ambiguous_candidates).
    The third element is non-empty only when detection was ambiguous.
    """
    warnings: list[str] = []
    grade_aliases = {normalize_text(a) for a in COLUMN_ALIASES["grade"]}

    # 1. Exact name match
    by_name = detect_column(list(df.columns), "grade")
    if by_name is not None:
        return by_name, warnings, []

    # 2. Prefix match: column starts with a grade alias + separator
    prefix_matches: list[str] = []
    for col in df.columns:
        norm = normalize_text(col)
        for alias in grade_aliases:
            if norm.startswith(alias) and len(norm) > len(alias):
                # Must be followed by a separator-like char (space, /)
                rest = norm[len(alias) :]
                if rest[0] in (" ", "/"):
                    prefix_matches.append(col)
                    break

    if len(prefix_matches) == 1:
        warnings.append(
            f"  Grade column detected by prefix match: '{prefix_matches[0]}'."
        )
        return prefix_matches[0], warnings, []
    elif len(prefix_matches) > 1:
        warnings.append(
            f"  AMBIGUOUS: multiple grade-like columns by prefix: {prefix_matches}."
        )
        return None, warnings, prefix_matches

    # 3. Content-based
    grade_tokens = {"ABS", "DEF", "ABJ", "ABI", ""}
    candidates: list[str] = []

    for col in df.columns:
        if col in known_columns:
            continue
        series = df[col].dropna().astype(str)
        if len(series) == 0:
            continue
        numeric_like = 0
        for val in series:
            cleaned = val.strip().replace(",", ".").upper()
            if cleaned in grade_tokens:
                numeric_like += 1
                continue
            # strip trailing "/xx" (e.g. "15/20")
            cleaned = re.sub(r"/\s*\d+(\.\d+)?$", "", cleaned)
            try:
                float(cleaned)
                numeric_like += 1
            except ValueError:
                pass
        if numeric_like / len(series) >= 0.5:
            candidates.append(col)

    if len(candidates) == 1:
        warnings.append(
            f"  Grade column auto-detected by content: '{candidates[0]}' "
            "(no column name match)."
        )
        return candidates[0], warnings, []
    elif len(candidates) > 1:
        # Heuristic: if there are sub-score columns (Q1, Q2, ...) and one
        # candidate looks like a summary column, prefer it.
        summary_aliases = {"total", "total general", "note totale", "somme", "sum"}
        summary_cols = [c for c in candidates if normalize_text(c) in summary_aliases]
        has_subscores = any(re.match(r"^[Qq]\d+$", c.strip()) for c in df.columns)

        if len(summary_cols) == 1 and has_subscores:
            warnings.append(
                f"  Grade column auto-detected as summary column: "
                f"'{summary_cols[0]}' (sub-score columns present)."
            )
            return summary_cols[0], warnings, []

        warnings.append(f"  AMBIGUOUS: multiple possible grade columns: {candidates}.")
        return None, warnings, candidates
    else:
        warnings.append("  No grade column detected. Skipping this file.")
        return None, warnings, []


# ============================================================================
# 2. FILE READING
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
    """Read a single sheet from an Excel file and apply standard cleaning."""
    df = clean_column_names(
        pd.read_excel(path, engine=engine, sheet_name=sheet, dtype=str)
    )
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
    (i.e. have at least 2 column alias matches).
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

    CSV: tries several encodings and auto-detects the delimiter via the
    Python csv sniffer (``sep=None, engine='python'``).

    XLSX/ODS: reads the first sheet only (use ``read_file_sheets`` for
    multi-sheet support).

    All formats: column names are cleaned of BOM markers and zero-width
    characters after reading.  If the first row doesn't look like headers
    (e.g. a TA put a title row), scans the first 10 rows for the real
    header row and promotes it.
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
                    pd.read_csv(
                        path, sep=None, engine="python", encoding=enc, dtype=str
                    )
                )
                break
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        if df is None:
            raise ValueError(
                f"Could not read '{path}' with any of the attempted encodings."
            )
        # Post-read check: if a non-UTF-7 encoding succeeded but column
        # names/data contain UTF-7 escape sequences, re-read with UTF-7.
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

    # Check if headers are in a later row (TA added title/metadata above)
    header_row = find_header_row(df)
    if header_row is not None:
        df = clean_column_names(promote_header_row(df, header_row))

    return df


# ============================================================================
# 3. NAME NORMALISATION & MATCHING
# ============================================================================


def normalize_name(name: str) -> str:
    """
    Normalise a student name for comparison:
    lowercase, strip accents, normalise hyphens → spaces, collapse whitespace.
    """
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
    """Return a human-readable hint about how similar two names are."""
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
    Interactively ask the user to confirm an ID match despite a name mismatch.

    Shows the edit distance and a similarity hint.
    Returns True to proceed with the match, False to skip.
    """
    norm_master = normalize_name(master_name)
    norm_ta = normalize_name(ta_name)
    dist = levenshtein_distance(norm_master, norm_ta)
    max_len = max(len(norm_master), len(norm_ta))
    hint = name_similarity_hint(dist, max_len)

    print(f"\n  Name mismatch for ID '{sid}' in '{filename}':")
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
            print("  -> Confirmed.")
            return True
        if raw in ("n", "no"):
            print("  -> Skipping this student.")
            return False
        print("  Please enter y or n.")


# ============================================================================
# 4. GRADE PARSING
# ============================================================================

# Tokens treated as "no grade" (absent, défaillant, …)
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
    """Result of parsing a raw grade string."""

    value: float | None = None  # None if absent / unparseable
    is_absent: bool = False
    raw: str = ""
    warning: str | None = None


def parse_grade(raw) -> ParsedGrade:
    """Parse a raw grade value into a numeric float or an absent marker."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ParsedGrade(raw="", warning="empty cell")

    raw_str = str(raw).strip()
    if raw_str == "":
        return ParsedGrade(raw=raw_str, warning="empty cell")

    upper = raw_str.upper()
    if upper in ABSENT_TOKENS:
        return ParsedGrade(is_absent=True, raw=raw_str)

    # Normalise: comma → dot, strip "/xx" suffix
    cleaned = raw_str.replace(",", ".")
    cleaned = re.sub(r"\s*/\s*\d+(\.\d+)?\s*$", "", cleaned)

    try:
        value = float(cleaned)
        return ParsedGrade(value=value, raw=raw_str)
    except ValueError:
        return ParsedGrade(
            raw=raw_str,
            warning=f"could not parse grade value '{raw_str}'",
        )


# ============================================================================
# 5. MASTER INDEX
# ============================================================================


@dataclass
class Student:
    """A student record from the master file."""

    student_id: str
    first_name: str
    last_name: str
    email: str
    grade: float | None = None
    grade_source: str | None = None  # which TA file set the grade
    is_absent: bool = False
    warnings: list[str] = field(default_factory=list)


@dataclass
class MasterIndex:
    """Indexes for fast student lookup."""

    by_id: dict[str, Student] = field(default_factory=dict)
    by_name: dict[str, list[Student]] = field(default_factory=dict)
    all_students: list[Student] = field(default_factory=list)


def build_master_index(df: pd.DataFrame) -> tuple[MasterIndex, list[str]]:
    """
    Build lookup indexes from the master DataFrame.
    Returns (MasterIndex, warnings).
    """
    warnings: list[str] = []
    cols = list(df.columns)

    id_col = detect_column(cols, "id")
    fn_col = detect_column(cols, "first_name")
    ln_col = detect_column(cols, "last_name")
    em_col = detect_column(cols, "email")

    missing = []
    if id_col is None:
        missing.append("student ID")
    if fn_col is None:
        missing.append("first name")
    if ln_col is None:
        missing.append("last name")
    if em_col is None:
        missing.append("email")
    if missing:
        raise ValueError(
            f"Master file is missing required columns: {', '.join(missing)}. "
            f"Available columns: {cols}"
        )

    index = MasterIndex()

    for _, row in df.iterrows():
        sid = str(row[id_col]).strip()
        fn = str(row[fn_col]).strip()
        ln = str(row[ln_col]).strip()
        email = str(row[em_col]).strip()

        student = Student(student_id=sid, first_name=fn, last_name=ln, email=email)
        index.all_students.append(student)

        # ID index
        if sid in index.by_id:
            warnings.append(
                f"  Duplicate student ID '{sid}' in master file. "
                "Keeping first occurrence."
            )
        else:
            index.by_id[sid] = student

        # Name index
        nk = make_name_key(fn, ln)
        index.by_name.setdefault(nk, []).append(student)

    return index, warnings


# ============================================================================
# 6. PROCESS A SINGLE TA FILE
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
    new_overrides: dict[str, str] = field(default_factory=dict)
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
    print(f"\n  Multiple {label} columns detected in '{filename}':")
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
            print(f"  -> Using '{chosen}'")
            return chosen
        elif choice == len(candidates) + 1:
            print("  -> Skipping file.")
            return None
        else:
            print(f"  Please enter a number between 1 and {len(candidates) + 1}.")


def _resolve_column(
    role: str,
    cols: list[str],
    overrides: dict[str, str],
    filename: str,
    interactive: bool,
    report: FileReport,
) -> tuple[str | None, dict[str, str]]:
    """
    Resolve a column for a given role, checking overrides first, then
    detecting, and prompting on ambiguity.

    Returns (column_name_or_None, new_overrides_to_save).
    """
    new_overrides: dict[str, str] = {}

    # 1. Check override
    if role in overrides:
        override_col = overrides[role]
        if override_col in cols:
            label = _ROLE_LABELS.get(role, role)
            report.warnings.append(
                f"  Using saved override for {label}: '{override_col}'."
            )
            return override_col, new_overrides
        report.warnings.append(
            f"  Saved override '{override_col}' for {role} not found in columns. "
            "Falling back to detection."
        )

    # 2. Detect
    matches = detect_all_columns(cols, role)

    if len(matches) == 1:
        return matches[0], new_overrides
    elif len(matches) == 0:
        return None, new_overrides
    else:
        # Ambiguous
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
        # Non-interactive: use first match
        report.warnings.append(
            f"  Using first match: '{matches[0]}' (non-interactive mode)."
        )
        return matches[0], new_overrides


def process_ta_file(
    path: str | Path,
    master: MasterIndex,
    *,
    interactive: bool = True,
    column_overrides: dict[str, str] | None = None,
    name_confirmations: list[str] | None = None,
) -> FileReport:
    """Process one TA grade file and update the master index in place.

    ``column_overrides`` is an optional dict of role -> column name that were
    previously saved from interactive choices.  New interactive choices are
    stored in ``report.new_overrides`` for the caller to persist.

    ``name_confirmations`` is an optional list of student IDs for which a
    name mismatch was previously confirmed.  New confirmations are stored
    in ``report.new_name_confirmations``.
    """
    path = Path(path)
    report = FileReport(filename=str(path))
    overrides = column_overrides or {}
    confirmed_ids = set(name_confirmations or [])

    # Read all sheets (CSV returns one sheet with name "")
    try:
        sheets = read_file_sheets(path)
    except Exception as e:
        report.warnings.append(f"  Could not read file: {e}")
        report.skipped = True
        return report

    if not sheets:
        report.warnings.append("  No valid sheets found in file.")
        report.skipped = True
        return report

    if len(sheets) > 1:
        sheet_names = [name for name, _ in sheets]
        report.warnings.append(
            f"  Found {len(sheets)} sheets with grade data: {sheet_names}."
        )

    all_new_overrides: dict[str, str] = {}
    # Working copy of overrides that accumulates interactive choices
    # so subsequent sheets benefit from earlier selections
    active_overrides = dict(overrides)

    for sheet_name, df in sheets:
        sheet_label = f" [sheet '{sheet_name}']" if sheet_name else ""

        if df.empty:
            report.warnings.append(f"  Sheet{sheet_label} is empty. Skipping.")
            continue

        cols = list(df.columns)

        # Detect columns with override + ambiguity support
        id_col, id_ov = _resolve_column(
            "id", cols, active_overrides, path.name, interactive, report
        )
        fn_col, fn_ov = _resolve_column(
            "first_name", cols, active_overrides, path.name, interactive, report
        )
        ln_col, ln_ov = _resolve_column(
            "last_name", cols, active_overrides, path.name, interactive, report
        )

        known_cols = {c for c in [id_col, fn_col, ln_col] if c is not None}

        # Grade column: check override first, then use the advanced detection
        grade_col: str | None = None
        grade_ov: dict[str, str] = {}

        if "grade" in active_overrides:
            override_grade = active_overrides["grade"]
            if override_grade in cols:
                report.warnings.append(
                    f"  Using saved override for grade: '{override_grade}'."
                )
                grade_col = override_grade
            else:
                report.warnings.append(
                    f"  Saved override '{override_grade}' for grade not found "
                    "in columns. Falling back to detection."
                )

        if grade_col is None:
            grade_col, gw, ambiguous = detect_grade_column(df, known_cols)
            report.warnings.extend(gw)

            if grade_col is None and ambiguous and interactive:
                grade_col = prompt_column_choice(ambiguous, path.name, "grade")
                if grade_col is not None:
                    report.warnings.append(
                        f"  Grade column manually selected: '{grade_col}'."
                    )
                    grade_ov["grade"] = grade_col

        if grade_col is None:
            report.warnings.append(
                f"  No grade column found{sheet_label}. Skipping sheet."
            )
            continue

        sheet_overrides = {**id_ov, **fn_ov, **ln_ov, **grade_ov}
        all_new_overrides.update(sheet_overrides)
        active_overrides.update(sheet_overrides)

        # Determine matching strategy
        use_id = id_col is not None
        use_name = fn_col is not None and ln_col is not None

        if not use_id and not use_name:
            report.warnings.append(
                f"  Sheet{sheet_label} has neither a student ID column nor "
                "both first/last name columns. Skipping sheet."
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
                            # Compare full concatenated names first — handles
                            # cases where first/last split differs between
                            # master and TA (e.g. "Mohamed Ayoub" / "Mebarki"
                            # vs "Mohamed" / "Ayoub Mebarki")
                            master_full_norm = f"{expected_fn} {expected_ln}"
                            ta_full_norm = f"{got_fn} {got_ln}"
                            if master_full_norm != ta_full_norm:
                                master_full = (
                                    f"{student.first_name} {student.last_name}"
                                )
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
                                            f"mismatch confirmed by user."
                                        )
                                    else:
                                        report.warnings.append(
                                            f"  Row {row_idx}: ID '{sid}' name "
                                            f"mismatch rejected by user. "
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
                        # ID not in master — try name fallback
                        if use_name:
                            fn_raw = str(row[fn_col]).strip()
                            ln_raw = str(row[ln_col]).strip()
                            report.warnings.append(
                                f"  Row {row_idx}: ID '{sid}' not found in master. "
                                f"Attempting name fallback ({fn_raw} {ln_raw})."
                            )
                        else:
                            report.warnings.append(
                                f"  Row {row_idx}: ID '{sid}' not found in master "
                                "and no name columns to fall back on. Skipping row."
                            )
                            continue

            # Fallback to name matching (or primary if no ID column)
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

                if len(matches) == 1:
                    student = matches[0]
                elif len(matches) > 1:
                    report.warnings.append(
                        f"  Row {row_idx}: name '{fn_raw} {ln_raw}' matches "
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
                    f"(ID={student.student_id}) already has a grade from '{prev_src}'. "
                    f"Duplicate found in '{path.name}'. Keeping first grade."
                )
                continue

            # Assign grade
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
            # else: empty cell, leave ungraded

    # Collect all new overrides for persistence
    report.new_overrides = all_new_overrides

    # If no sheets yielded any matches, mark as skipped
    if report.students_matched == 0 and report.grades_assigned == 0:
        report.skipped = True

    return report


# ============================================================================
# 7. OUTPUT
# ============================================================================


def write_moodle_csv(master: MasterIndex, output_path: str | Path) -> None:
    """
    Write a Moodle-compatible CSV.
    Columns: Identifier, Email address, First name, Last name, Grade
    """
    output_path = Path(output_path)
    rows = []
    for s in master.all_students:
        grade_str = ""
        if s.is_absent:
            grade_str = "ABS"
        elif s.grade is not None:
            # Format: no unnecessary decimals
            grade_str = f"{s.grade:g}"
        rows.append(
            {
                "Identifier": s.student_id,
                "Email address": s.email,
                "First name": s.first_name,
                "Last name": s.last_name,
                "Grade": grade_str,
            }
        )

    df = pd.DataFrame(rows)
    df.to_csv(output_path, index=False, encoding="utf-8")


# ============================================================================
# 8. SUMMARY
# ============================================================================


def print_summary(
    master: MasterIndex,
    reports: list[FileReport],
    output_path: str,
) -> None:
    """Print a human-readable summary of the consolidation."""
    total_students = len(master.all_students)
    graded = sum(1 for s in master.all_students if s.grade is not None)
    absent = sum(1 for s in master.all_students if s.is_absent)
    no_grade = total_students - graded - absent

    files_ok = sum(1 for r in reports if not r.skipped)
    files_skipped = sum(1 for r in reports if r.skipped)

    print("\n" + "=" * 60)
    print("CONSOLIDATION SUMMARY")
    print("=" * 60)
    print(f"  Master roster          : {total_students} students")
    print(f"  TA files processed     : {files_ok}")
    print(f"  TA files skipped       : {files_skipped}")
    print(f"  Students with grade    : {graded}")
    print(f"  Students absent (ABS…) : {absent}")
    print(f"  Students without grade : {no_grade}")
    print(f"  Output file            : {output_path}")

    # Per-file details
    for r in reports:
        status = "SKIPPED" if r.skipped else "OK"
        print(f"\n  [{status}] {r.filename}")
        if not r.skipped:
            print(
                f"    matched={r.students_matched}  "
                f"grades={r.grades_assigned}  "
                f"absent={r.students_absent}"
            )
        for w in r.warnings:
            print(f"    {w}")

    # List students without a grade
    missing = [s for s in master.all_students if s.grade is None and not s.is_absent]
    if missing:
        print(f"\n  Students without any grade ({len(missing)}):")
        for s in missing:
            print(f"    - {s.last_name}, {s.first_name} (ID={s.student_id})")

    print("=" * 60 + "\n")


# ============================================================================
# 9. MAIN
# ============================================================================


SUPPORTED_EXTENSIONS = {".csv", ".tsv", ".txt", ".xlsx", ".ods"}


def resolve_grade_files(
    config_dir: Path,
    grade_files: list[str] | None = None,
    grade_dir: str | None = None,
) -> list[Path]:
    """
    Build the list of grade file paths from explicit file list, directory scan,
    or both.  Deduplicates by resolved absolute path and sorts for determinism.
    """
    seen: set[Path] = set()
    result: list[Path] = []

    def _add(p: Path) -> None:
        resolved = p.resolve()
        if resolved not in seen:
            seen.add(resolved)
            result.append(p)

    # Explicit file list
    if grade_files:
        for gf in grade_files:
            gf_path = Path(gf)
            if not gf_path.is_absolute():
                gf_path = config_dir / gf_path
            _add(gf_path)

    # Directory scan
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
        raise ValueError(
            "Config must specify 'grade_files' (list) and/or 'grade_dir' (path)."
        )

    cfg.setdefault("output_file", "grades_consolidated.csv")
    return cfg


def save_config(config_path: str | Path, cfg: dict) -> None:
    """Write the config dict back to the YAML file, preserving new overrides."""
    config_path = Path(config_path)
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(cfg, f, default_flow_style=False, allow_unicode=True, sort_keys=False)


def consolidate(
    config_path: str | Path, *, interactive: bool = True
) -> tuple[MasterIndex, list[FileReport]]:
    """
    Run the full consolidation pipeline.
    Returns (master_index, list_of_file_reports) for programmatic use.
    """
    cfg = load_config(config_path)

    # Resolve paths relative to the config file location
    config_dir = Path(config_path).parent

    master_path = Path(cfg["master_file"])
    if not master_path.is_absolute():
        master_path = config_dir / master_path

    output_path = Path(cfg["output_file"])
    if not output_path.is_absolute():
        output_path = config_dir / output_path

    # Read master
    print(f"Reading master file: {master_path}")
    master_df = read_file(master_path)
    master, master_warnings = build_master_index(master_df)
    if master_warnings:
        print("Master file warnings:")
        for w in master_warnings:
            print(w)

    # Resolve grade file list
    grade_file_paths = resolve_grade_files(
        config_dir,
        grade_files=cfg.get("grade_files"),
        grade_dir=cfg.get("grade_dir"),
    )

    if not grade_file_paths:
        print("WARNING: No grade files found. Output will have no grades.")

    # Load existing column overrides and name confirmations
    all_overrides: dict[str, dict[str, str]] = cfg.get("column_overrides", {})
    all_name_confirmations: dict[str, list[str]] = cfg.get("name_confirmations", {})
    config_changed = False

    # Process TA files
    reports: list[FileReport] = []
    for gf_path in grade_file_paths:
        print(f"Processing: {gf_path}")
        file_overrides = all_overrides.get(gf_path.name, {})
        file_name_confs = all_name_confirmations.get(gf_path.name, [])
        report = process_ta_file(
            gf_path,
            master,
            interactive=interactive,
            column_overrides=file_overrides,
            name_confirmations=file_name_confs,
        )
        reports.append(report)

        # Collect new column overrides from interactive choices
        if report.new_overrides:
            existing = all_overrides.get(gf_path.name, {})
            existing.update(report.new_overrides)
            all_overrides[gf_path.name] = existing
            config_changed = True

        # Collect new name confirmations
        if report.new_name_confirmations:
            existing_confs = all_name_confirmations.get(gf_path.name, [])
            existing_confs.extend(report.new_name_confirmations)
            all_name_confirmations[gf_path.name] = existing_confs
            config_changed = True

    # Persist overrides back to config if any new choices were made
    if config_changed:
        if all_overrides:
            cfg["column_overrides"] = all_overrides
        if all_name_confirmations:
            cfg["name_confirmations"] = all_name_confirmations
        save_config(config_path, cfg)
        print(f"\n  Overrides saved to {config_path}")

    # Write output
    write_moodle_csv(master, output_path)
    print_summary(master, reports, str(output_path))

    return master, reports


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Consolidate student grades into a Moodle-compatible CSV.",
    )
    parser.add_argument(
        "config",
        help="Path to the YAML configuration file.",
    )
    args = parser.parse_args()
    consolidate(args.config)


if __name__ == "__main__":
    main()
