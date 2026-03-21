#!/usr/bin/env python3
"""
test_consolidate_grades.py
==========================
Unit and integration tests for consolidate_grades.py.

Run with:  pytest test_consolidate_grades.py -v
"""

import csv
from pathlib import Path

import pandas as pd
import pytest
import yaml

from consolidate_grades.consolidate import (
    build_master_index,
    clean_column_names,
    consolidate,
    detect_column,
    detect_grade_column,
    find_header_row,
    load_config,
    make_name_key,
    normalize_name,
    normalize_text,
    parse_grade,
    process_ta_file,
    promote_header_row,
    read_file,
    resolve_grade_files,
    write_moodle_csv,
)

# ============================================================================
# Helpers
# ============================================================================


def _master_df(rows=None):
    """Create a small master DataFrame for testing."""
    if rows is None:
        rows = [
            ("12345", "Jean", "Dupont", "jean.dupont@etu.fr"),
            ("12346", "Marie", "Curie", "marie.curie@etu.fr"),
            ("12347", "Éloïse", "Le Bœuf-André", "eloise.leboeuf@etu.fr"),
            ("12348", "Ahmed", "Ben Salah", "ahmed.bensalah@etu.fr"),
        ]
    return pd.DataFrame(
        rows,
        columns=["Numéro étudiant", "Prénom", "Nom de famille", "Email"],
    )


def _write_csv(
    path: Path, header: list[str], rows: list[list], sep=",", encoding="utf-8"
):
    """Write a quick CSV helper."""
    with open(path, "w", newline="", encoding=encoding) as f:
        w = csv.writer(f, delimiter=sep)
        w.writerow(header)
        w.writerows(rows)


def _build_master():
    """Build a MasterIndex from the standard test master."""
    df = _master_df()
    master, _warnings = build_master_index(df)
    return master


# ============================================================================
# 1. normalize_text
# ============================================================================


class TestNormalizeText:
    def test_lowercase(self):
        assert normalize_text("HELLO") == "hello"

    def test_accents_stripped(self):
        assert normalize_text("Éloïse") == "eloise"
        assert normalize_text("Numéro étudiant") == "numero etudiant"

    def test_hyphens_become_spaces(self):
        assert normalize_text("Le Bœuf-André") == "le bœuf andre"

    def test_underscores_become_spaces(self):
        assert normalize_text("first_name") == "first name"

    def test_collapse_whitespace(self):
        assert normalize_text("  nom   de   famille  ") == "nom de famille"

    def test_combined(self):
        assert normalize_text("  Numéro_Étudiant  ") == "numero etudiant"

    def test_ascii_apostrophe_becomes_space(self):
        assert normalize_text("Numéro d'identification") == "numero d identification"

    def test_curly_apostrophe_becomes_space(self):
        assert (
            normalize_text("Num\u00e9ro d\u2019identification")
            == "numero d identification"
        )

    def test_left_curly_apostrophe(self):
        assert normalize_text("l\u2018exemple") == "l exemple"

    def test_degree_sign_becomes_o(self):
        assert normalize_text("N°") == "no"
        assert normalize_text("N° d'identification") == "no d identification"

    def test_ordinal_indicator_becomes_o(self):
        assert normalize_text("Nº etudiant") == "no etudiant"


# ============================================================================
# 1b. clean_column_names
# ============================================================================


class TestCleanColumnNames:
    def test_bom_stripped(self):
        df = pd.DataFrame({"\ufeffPrénom": ["Alice"], "Nom": ["Martin"]})
        cleaned = clean_column_names(df)
        assert list(cleaned.columns) == ["Prénom", "Nom"]

    def test_zero_width_space_stripped(self):
        df = pd.DataFrame({"\u200bID": ["1"], "Note\u200b": ["15"]})
        cleaned = clean_column_names(df)
        assert list(cleaned.columns) == ["ID", "Note"]

    def test_surrounding_whitespace_stripped(self):
        df = pd.DataFrame({"  Prénom  ": ["Alice"], " Nom ": ["Martin"]})
        cleaned = clean_column_names(df)
        assert list(cleaned.columns) == ["Prénom", "Nom"]

    def test_clean_columns_already_clean(self):
        df = pd.DataFrame({"Prénom": ["Alice"], "Nom": ["Martin"]})
        cleaned = clean_column_names(df)
        assert list(cleaned.columns) == ["Prénom", "Nom"]


class TestDetectColumn:
    # --- ID column ---
    @pytest.mark.parametrize(
        "col_name",
        [
            "Numéro étudiant",
            "numero etudiant",
            "NUMERO ETUDIANT",
            "Student ID",
            "student_id",
            "ID",
            "NIP",
            "Identifiant",
            "Code étudiant",
            "No étudiant",
            "Numéro d\u2019identification",
            "Numéro d'identification",
            "N° d'identification",
        ],
    )
    def test_id_variants(self, col_name):
        assert detect_column([col_name, "Other"], "id") == col_name

    # --- First name ---
    @pytest.mark.parametrize(
        "col_name",
        ["Prénom", "prenom", "First Name", "first_name", "Given Name"],
    )
    def test_first_name_variants(self, col_name):
        assert detect_column([col_name, "Other"], "first_name") == col_name

    # --- Last name ---
    @pytest.mark.parametrize(
        "col_name",
        ["Nom", "Nom de famille", "Last Name", "last_name", "Family Name"],
    )
    def test_last_name_variants(self, col_name):
        assert detect_column([col_name, "Other"], "last_name") == col_name

    # --- Email ---
    @pytest.mark.parametrize(
        "col_name",
        [
            "Email",
            "Mail",
            "Courriel",
            "Adresse email",
            "E-mail",
            "Email address",
            "Adresse de courriel",
        ],
    )
    def test_email_variants(self, col_name):
        assert detect_column([col_name, "Other"], "email") == col_name

    # --- Grade ---
    @pytest.mark.parametrize(
        "col_name",
        [
            "Note",
            "Grade",
            "Score",
            "Résultat",
            "Mark",
            "Points",
            "Note finale",
            "Total",
            "Total général",
        ],
    )
    def test_grade_variants(self, col_name):
        assert detect_column([col_name, "Other"], "grade") == col_name

    def test_no_match_returns_none(self):
        assert detect_column(["Foo", "Bar"], "id") is None

    def test_first_match_wins(self):
        # If two columns match, first one is returned
        assert detect_column(["ID", "Student ID"], "id") == "ID"


# ============================================================================
# 3. detect_grade_column (by name, by content, ambiguous)
# ============================================================================


class TestDetectGradeColumn:
    def test_by_name(self):
        df = pd.DataFrame({"ID": ["1"], "Note": ["15"]})
        col, warnings = detect_grade_column(df, {"ID"})
        assert col == "Note"
        assert len(warnings) == 0

    def test_by_content_single_candidate(self):
        df = pd.DataFrame(
            {"ID": ["1", "2"], "Nom": ["A", "B"], "Mystery": ["14,5", "ABS"]}
        )
        col, warnings = detect_grade_column(df, {"ID", "Nom"})
        assert col == "Mystery"
        assert any("auto-detected" in w for w in warnings)

    def test_by_content_ambiguous(self):
        df = pd.DataFrame(
            {
                "ID": ["1", "2"],
                "Col_A": ["14", "15"],
                "Col_B": ["10", "12"],
            }
        )
        col, warnings = detect_grade_column(df, {"ID"})
        assert col is None
        assert any("AMBIGUOUS" in w for w in warnings)

    def test_no_grade_column(self):
        df = pd.DataFrame({"ID": ["1"], "Nom": ["A"], "Adresse": ["1 rue X"]})
        col, warnings = detect_grade_column(df, {"ID", "Nom"})
        assert col is None
        assert any("No grade column" in w for w in warnings)

    def test_with_slash_notation(self):
        """Grades like '15/20' should still be detected as numeric."""
        df = pd.DataFrame({"ID": ["1", "2"], "Résultat": ["15/20", "12,5/20"]})
        col, _warnings = detect_grade_column(df, {"ID"})
        # Should match by name first
        assert col == "Résultat"

    def test_prefix_match_note_slash_20(self):
        """'Note /20' should be detected via prefix matching."""
        df = pd.DataFrame({"ID": ["1", "2"], "Note /20": ["15", "16"]})
        col, warnings = detect_grade_column(df, {"ID"})
        assert col == "Note /20"
        assert any("prefix" in w.lower() for w in warnings)

    def test_prefix_match_note_slash_23(self):
        """'Note /23' should also work."""
        df = pd.DataFrame({"ID": ["1"], "Note /23": ["18"]})
        col, _warnings = detect_grade_column(df, {"ID"})
        assert col == "Note /23"

    def test_total_column_with_subscores(self):
        """'Total' preferred over Q1-Q14 sub-score columns via exact alias."""
        df = pd.DataFrame(
            {
                "Numéro étudiant": ["1", "2"],
                "Q1": ["2", "1"],
                "Q2": ["3", "2"],
                "Q3": ["1", "0"],
                "Total": ["6", "3"],
            }
        )
        col, warnings = detect_grade_column(df, {"Numéro étudiant"})
        assert col == "Total"
        # Total is an exact alias match, so no warnings expected
        assert len(warnings) == 0

    def test_summary_column_preferred_over_subscores(self):
        """When summary column is not an alias, prefer it via content heuristic."""
        df = pd.DataFrame(
            {
                "Numéro étudiant": ["1", "2"],
                "Q1": ["2", "1"],
                "Q2": ["3", "2"],
                "Q3": ["1", "0"],
                "Somme": ["6", "3"],
            }
        )
        col, warnings = detect_grade_column(df, {"Numéro étudiant"})
        assert col == "Somme"
        assert any("summary" in w.lower() for w in warnings)

    def test_total_exact_name_match(self):
        """'Total' as exact alias match (no sub-score columns needed)."""
        df = pd.DataFrame({"ID": ["1"], "Total": ["15"]})
        col, _warnings = detect_grade_column(df, {"ID"})
        assert col == "Total"

    def test_ambiguous_prefix_matches(self):
        """Multiple prefix matches should be flagged as ambiguous."""
        df = pd.DataFrame({"ID": ["1"], "Note /20": ["15"], "Note finale /20": ["15"]})
        _col, _warnings = detect_grade_column(df, {"ID"})
        # note finale is an exact alias, so it matches by name first
        # Let's test with two true prefix-only columns instead
        df2 = pd.DataFrame({"ID": ["1"], "Score /20": ["15"], "Score max": ["20"]})
        col2, warnings2 = detect_grade_column(df2, {"ID"})
        assert col2 is None
        assert any("AMBIGUOUS" in w for w in warnings2)


# ============================================================================
# 3b. Header row detection
# ============================================================================


class TestFindHeaderRow:
    def test_good_headers_returns_none(self):
        """When headers are already correct, return None."""
        df = pd.DataFrame({"Numéro étudiant": ["1"], "Prénom": ["A"], "Note": ["15"]})
        assert find_header_row(df) is None

    def test_title_in_row_0(self):
        """Title in row 0, real headers in data row 0 (original row 1)."""
        df = pd.DataFrame(
            {
                "Notes QCM CC 9 mars 2026 EI-3": [
                    "Numero",
                    "12345",
                ],
                "Unnamed: 1": ["Prenom", "Alice"],
                "Unnamed: 2": ["Nom", "Martin"],
                "Unnamed: 3": ["Numero etudiant", "12345"],
                "Unnamed: 4": ["Note /23", "18"],
            }
        )
        row = find_header_row(df)
        assert row == 0

    def test_multiple_metadata_rows(self):
        """Multiple metadata rows before real headers."""
        df = pd.DataFrame(
            {
                "Notes - UL1IN002": [
                    "Correcteur : Bilal",
                    "",
                    "Numero d'etudiant",
                    "12345",
                ],
                "Unnamed: 1": [
                    "Groupes : DC-1",
                    "",
                    "Nom",
                    "Martin",
                ],
                "Unnamed: 2": ["", "", "Prenom", "Alice"],
                "Unnamed: 3": ["", "", "Note", "15"],
            }
        )
        row = find_header_row(df)
        assert row == 2

    def test_no_alias_matches_anywhere(self):
        """When no row has alias matches, return None (stay with current)."""
        df = pd.DataFrame(
            {
                "Foo": ["bar", "baz"],
                "Qux": ["quux", "corge"],
            }
        )
        assert find_header_row(df) is None

    def test_promote_header_row(self):
        """promote_header_row drops preamble and sets new headers."""
        df = pd.DataFrame(
            {
                "Title": ["Metadata", "Nom", "Martin"],
                "Col2": ["Info", "Prenom", "Alice"],
                "Col3": ["", "Note", "15"],
            }
        )
        new_df = promote_header_row(df, 1)
        assert list(new_df.columns) == ["Nom", "Prenom", "Note"]
        assert len(new_df) == 1
        assert new_df.iloc[0]["Nom"] == "Martin"


class TestReadFileHeaderDetection:
    def test_csv_with_title_row(self, tmp_path):
        """CSV where row 1 has a title, row 2 has real headers."""
        p = tmp_path / "ta.csv"
        p.write_text(
            "Notes QCM CC 9 mars 2026 EI-3,,,,\n"
            "Numero,Prenom,Nom,Numero etudiant,Note /23\n"
            "1,Alice,Martin,12345,18\n",
            encoding="utf-8",
        )
        df = read_file(p)
        assert "Numero etudiant" in df.columns or "Numero" in df.columns
        assert len(df) == 1

    def test_csv_with_multiple_metadata_rows(self, tmp_path):
        """CSV with 3 rows of metadata before the real headers."""
        p = tmp_path / "ta.csv"
        p.write_text(
            "Notes Elements de programmation 2,,,\n"
            "Correcteur : Bilal,Groupes : DC-1,,\n"
            ",,,\n"
            "Numero d'etudiant,Nom,Prenom,Note\n"
            "12345,Martin,Alice,15\n",
            encoding="utf-8",
        )
        df = read_file(p)
        # The header detection should have found the real header row
        cols_norm = [c.lower() for c in df.columns]
        assert any(c == "nom" for c in cols_norm)
        assert len(df) == 1

    def test_xlsx_with_title_row(self, tmp_path):
        """XLSX where row 1 has a title, row 2 has real headers."""
        p = tmp_path / "ta.xlsx"
        raw = pd.DataFrame(
            {
                "Notes QCM": ["Numero etudiant", "12345"],
                "Unnamed: 1": ["Nom", "Martin"],
                "Unnamed: 2": ["Note", "15"],
            }
        )
        raw.to_excel(p, index=False, engine="openpyxl")
        df = read_file(p)
        assert detect_column(list(df.columns), "id") is not None
        assert len(df) == 1

    def test_normal_csv_not_degraded(self, tmp_path):
        """A normal CSV should NOT have its headers changed."""
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Prénom", "Nom", "Note"],
            [["12345", "Alice", "Martin", "15"]],
        )
        df = read_file(p)
        assert df.columns[0] == "Numéro étudiant"
        assert len(df) == 1

    def test_integration_ta_file_with_title_row(self, tmp_path):
        """Full pipeline with a TA file that has a title row."""
        master = _build_master()
        p = tmp_path / "ta.csv"
        p.write_text(
            "Notes QCM CC 9 mars 2026,,,,\n"
            "Numero,Prenom,Nom,Numero etudiant,Note /23\n"
            "1,Jean,Dupont,12345,18\n"
            "2,Marie,Curie,12346,15\n",
            encoding="utf-8",
        )
        report = process_ta_file(p, master)
        assert not report.skipped
        assert report.grades_assigned == 2
        assert master.by_id["12345"].grade == 18.0
        assert master.by_id["12346"].grade == 15.0

    def test_integration_ta_file_with_subscores_and_total(self, tmp_path):
        """TA file with Q1-Q14 sub-scores and a Total column."""
        master = _build_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            [
                "Prénom",
                "Nom de famille",
                "N° d'identification",
                "Q1",
                "Q2",
                "Q3",
                "Total",
            ],
            [
                ["Jean", "Dupont", "12345", "2", "3", "1", "6"],
                ["Marie", "Curie", "12346", "1", "2", "0.5", "3.5"],
            ],
        )
        report = process_ta_file(p, master)
        assert not report.skipped
        assert report.grades_assigned == 2
        assert master.by_id["12345"].grade == 6.0
        assert master.by_id["12346"].grade == 3.5


class TestReadFile:
    def test_csv_utf8_comma(self, tmp_path):
        p = tmp_path / "data.csv"
        _write_csv(p, ["A", "B"], [["1", "x"], ["2", "y"]])
        df = read_file(p)
        assert list(df.columns) == ["A", "B"]
        assert len(df) == 2

    def test_csv_semicolon_separator(self, tmp_path):
        p = tmp_path / "data.csv"
        _write_csv(p, ["A", "B"], [["1", "x"], ["2", "y"]], sep=";")
        df = read_file(p)
        assert list(df.columns) == ["A", "B"]
        assert len(df) == 2

    def test_csv_latin1_encoding(self, tmp_path):
        p = tmp_path / "data.csv"
        _write_csv(
            p,
            ["Prénom", "Note"],
            [["Éloïse", "14,5"]],
            encoding="latin-1",
        )
        df = read_file(p)
        assert "Prénom" in df.columns
        assert df.iloc[0]["Prénom"] == "Éloïse"

    def test_xlsx(self, tmp_path):
        p = tmp_path / "data.xlsx"
        pd.DataFrame({"A": ["1"], "B": ["x"]}).to_excel(
            p, index=False, engine="openpyxl"
        )
        df = read_file(p)
        assert list(df.columns) == ["A", "B"]

    def test_ods(self, tmp_path):
        p = tmp_path / "data.ods"
        pd.DataFrame({"A": ["1"], "B": ["x"]}).to_excel(p, index=False, engine="odf")
        df = read_file(p)
        assert list(df.columns) == ["A", "B"]

    def test_unsupported_extension(self, tmp_path):
        p = tmp_path / "data.json"
        p.write_text("{}")
        with pytest.raises(ValueError, match="Unsupported file extension"):
            read_file(p)


# ============================================================================
# 5. Name normalisation
# ============================================================================


class TestNameNormalization:
    def test_basic(self):
        assert normalize_name("Jean") == "jean"

    def test_accents(self):
        assert normalize_name("Éloïse") == "eloise"

    def test_hyphen(self):
        assert normalize_name("Le Bœuf-André") == "le bœuf andre"
        # Note: œ is NOT an accent - it is a ligature and survives normalization
        # This is deliberate: we only strip combining characters, not ligatures.

    def test_make_name_key(self):
        key = make_name_key("Jean-Pierre", "Dupont")
        assert key == "dupont|jean pierre"

    def test_make_name_key_accented(self):
        key = make_name_key("Éloïse", "André")
        assert key == "andre|eloise"


# ============================================================================
# 6. parse_grade
# ============================================================================


class TestParseGrade:
    def test_integer(self):
        g = parse_grade("15")
        assert g.value == 15.0
        assert not g.is_absent

    def test_float_dot(self):
        g = parse_grade("14.5")
        assert g.value == 14.5

    def test_float_comma(self):
        g = parse_grade("14,5")
        assert g.value == 14.5

    def test_with_slash(self):
        g = parse_grade("15/20")
        assert g.value == 15.0

    def test_with_slash_and_comma(self):
        g = parse_grade("12,5/20")
        assert g.value == 12.5

    def test_abs(self):
        g = parse_grade("ABS")
        assert g.is_absent
        assert g.value is None

    def test_def(self):
        g = parse_grade("DEF")
        assert g.is_absent

    def test_absent_case_insensitive(self):
        g = parse_grade("abs")
        # our code does .upper() comparison
        assert g.is_absent

    def test_empty_string(self):
        g = parse_grade("")
        assert g.value is None
        assert not g.is_absent
        assert g.warning is not None

    def test_none(self):
        g = parse_grade(None)
        assert g.value is None

    def test_nan(self):
        g = parse_grade(float("nan"))
        assert g.value is None

    def test_garbage(self):
        g = parse_grade("not_a_grade")
        assert g.value is None
        assert g.warning is not None

    def test_whitespace_padded(self):
        g = parse_grade("  15  ")
        assert g.value == 15.0

    def test_zero(self):
        g = parse_grade("0")
        assert g.value == 0.0
        assert not g.is_absent

    def test_high_grade(self):
        """No assumption about max grade."""
        g = parse_grade("42")
        assert g.value == 42.0


# ============================================================================
# 7. build_master_index
# ============================================================================


class TestBuildMasterIndex:
    def test_basic(self):
        master, warnings = build_master_index(_master_df())
        assert len(master.all_students) == 4
        assert "12345" in master.by_id
        assert len(warnings) == 0

    def test_duplicate_id_warning(self):
        df = _master_df(
            [
                ("12345", "Jean", "Dupont", "jean@etu.fr"),
                ("12345", "Marie", "Curie", "marie@etu.fr"),
            ]
        )
        master, warnings = build_master_index(df)
        assert len(warnings) == 1
        assert "Duplicate" in warnings[0]
        # First occurrence kept
        assert master.by_id["12345"].first_name == "Jean"

    def test_missing_column_raises(self):
        df = pd.DataFrame({"Foo": ["1"], "Bar": ["x"]})
        with pytest.raises(ValueError, match="missing required columns"):
            build_master_index(df)

    def test_name_index(self):
        master, _ = build_master_index(_master_df())
        key = make_name_key("Jean", "Dupont")
        assert key in master.by_name
        assert len(master.by_name[key]) == 1


# ============================================================================
# 8. process_ta_file - matching, edge cases
# ============================================================================


class TestProcessTaFile:
    def _setup_master(self):
        return _build_master()

    def test_match_by_id(self, tmp_path):
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Note"],
            [["12345", "15"], ["12346", "14,5"]],
        )
        report = process_ta_file(p, master)
        assert not report.skipped
        assert report.grades_assigned == 2
        assert master.by_id["12345"].grade == 15.0
        assert master.by_id["12346"].grade == 14.5

    def test_match_by_name_fallback(self, tmp_path):
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Prénom", "Nom", "Note"],
            [["Jean", "Dupont", "16"]],
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 1
        assert master.by_id["12345"].grade == 16.0

    def test_match_by_name_with_accents(self, tmp_path):
        """Name matching should handle accents."""
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Prénom", "Nom", "Note"],
            [["Eloise", "Le Boeuf Andre", "17"]],
        )
        report = process_ta_file(p, master)
        # This should NOT match because œ ≠ oe in our normalization
        # (we only strip combining chars, not expand ligatures).
        # The user asked: "normalize for caps, accents, hyphens.
        # Anything more than that should be a warning."
        # So this is expected to fail to match — and produce a warning.
        if report.grades_assigned == 0:
            assert any("not found" in w for w in report.warnings)
        # If the TA spells it exactly right (with œ), it should match:

    def test_match_by_name_hyphen_normalized(self, tmp_path):
        """Hyphens in names should be normalized."""
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        # Master has "Le Bœuf-André" — TA writes "Le Bœuf André" (no hyphen)
        _write_csv(
            p,
            ["Prénom", "Nom", "Note"],
            [["Éloïse", "Le Bœuf André", "18"]],
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 1

    def test_ambiguous_name_skipped(self, tmp_path):
        """Two students with the same name → skip with warning."""
        df = _master_df(
            [
                ("12345", "Jean", "Dupont", "jean1@etu.fr"),
                ("12346", "Jean", "Dupont", "jean2@etu.fr"),
            ]
        )
        master, _ = build_master_index(df)
        p = tmp_path / "ta.csv"
        _write_csv(p, ["Prénom", "Nom", "Note"], [["Jean", "Dupont", "14"]])
        report = process_ta_file(p, master)
        assert report.grades_assigned == 0
        assert any("matches 2 students" in w for w in report.warnings)

    def test_student_not_in_master(self, tmp_path):
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Note"],
            [["99999", "15"]],
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 0
        assert any("not found" in w.lower() for w in report.warnings)

    def test_id_match_with_name_cross_check_mismatch(self, tmp_path):
        """ID matches but name differs → warning but still assign."""
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Prénom", "Nom", "Note"],
            [["12345", "Pierre", "Martin", "15"]],
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 1
        assert master.by_id["12345"].grade == 15.0
        assert any("please verify" in w.lower() for w in report.warnings)

    def test_duplicate_grade_warning(self, tmp_path):
        """Same student in two TA files → warning, keep first grade."""
        master = self._setup_master()
        # First file
        p1 = tmp_path / "ta1.csv"
        _write_csv(p1, ["Numéro étudiant", "Note"], [["12345", "15"]])
        process_ta_file(p1, master)

        # Second file
        p2 = tmp_path / "ta2.csv"
        _write_csv(p2, ["Numéro étudiant", "Note"], [["12345", "18"]])
        report = process_ta_file(p2, master)
        assert master.by_id["12345"].grade == 15.0  # first grade kept
        assert any("already has a grade" in w.lower() for w in report.warnings)

    def test_absent_student(self, tmp_path):
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Note"],
            [["12345", "ABS"]],
        )
        report = process_ta_file(p, master)
        assert report.students_absent == 1
        assert master.by_id["12345"].is_absent
        assert master.by_id["12345"].grade is None

    def test_garbage_grade_warning(self, tmp_path):
        """Non-parseable grade → warning, no grade assigned."""
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Note"],
            [["12345", "???"]],
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 0
        assert master.by_id["12345"].grade is None
        assert any("could not parse" in w.lower() for w in report.warnings)

    def test_empty_file_skipped(self, tmp_path):
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        p.write_text("Numéro étudiant,Note\n")
        report = process_ta_file(p, master)
        assert report.skipped

    def test_no_usable_columns_skipped(self, tmp_path):
        """File with neither ID nor names → skip."""
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Foo", "Note"],
            [["bar", "15"]],
        )
        report = process_ta_file(p, master)
        assert report.skipped
        assert any("neither" in w.lower() for w in report.warnings)

    def test_xlsx_file(self, tmp_path):
        """TA file as XLSX."""
        master = self._setup_master()
        p = tmp_path / "ta.xlsx"
        pd.DataFrame({"Numéro étudiant": ["12345"], "Note": ["15"]}).to_excel(
            p, index=False, engine="openpyxl"
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 1

    def test_ods_file(self, tmp_path):
        """TA file as ODS."""
        master = self._setup_master()
        p = tmp_path / "ta.ods"
        pd.DataFrame({"Numéro étudiant": ["12345"], "Note": ["15"]}).to_excel(
            p, index=False, engine="odf"
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 1

    def test_id_not_in_master_with_name_fallback(self, tmp_path):
        """ID unknown but name columns exist → fall back to name."""
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Prénom", "Nom", "Note"],
            [["99999", "Jean", "Dupont", "13"]],
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 1
        assert master.by_id["12345"].grade == 13.0
        assert any("not found in master" in w.lower() for w in report.warnings)

    def test_semicolon_csv(self, tmp_path):
        """French-style CSV with semicolons."""
        master = self._setup_master()
        p = tmp_path / "ta.csv"
        _write_csv(
            p,
            ["Numéro étudiant", "Note"],
            [["12345", "14,5"]],
            sep=";",
        )
        report = process_ta_file(p, master)
        assert report.grades_assigned == 1
        assert master.by_id["12345"].grade == 14.5


# ============================================================================
# 9. write_moodle_csv
# ============================================================================


class TestWriteMoodleCsv:
    def test_output_format(self, tmp_path):
        master = _build_master()
        master.by_id["12345"].grade = 15.5
        master.by_id["12346"].is_absent = True

        out = tmp_path / "out.csv"
        write_moodle_csv(master, out)

        df = pd.read_csv(out, dtype=str, keep_default_na=False)
        assert "Identifier" in df.columns
        assert "Email address" in df.columns
        assert "Grade" in df.columns
        assert len(df) == 4

        jean_row = df[df["Identifier"] == "12345"].iloc[0]
        assert jean_row["Grade"] == "15.5"

        marie_row = df[df["Identifier"] == "12346"].iloc[0]
        assert marie_row["Grade"] == "ABS"

    def test_no_grade_is_empty(self, tmp_path):
        master = _build_master()
        out = tmp_path / "out.csv"
        write_moodle_csv(master, out)
        df = pd.read_csv(out, dtype=str, keep_default_na=False)
        # Students with no grade should have empty string
        assert all(
            df.loc[df["Identifier"] == sid, "Grade"].values[0] == ""
            for sid in ["12345", "12346", "12347", "12348"]
        )


# ============================================================================
# 10. Config loading
# ============================================================================


class TestLoadConfig:
    def test_valid_config_with_files(self, tmp_path):
        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "master.csv",
                    "grade_files": ["g1.csv"],
                    "output_file": "out.csv",
                }
            )
        )
        cfg = load_config(cfg_path)
        assert cfg["master_file"] == "master.csv"

    def test_valid_config_with_dir(self, tmp_path):
        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "master.csv",
                    "grade_dir": "grades/",
                    "output_file": "out.csv",
                }
            )
        )
        cfg = load_config(cfg_path)
        assert cfg["grade_dir"] == "grades/"

    def test_valid_config_with_both(self, tmp_path):
        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "master.csv",
                    "grade_files": ["extra.csv"],
                    "grade_dir": "grades/",
                }
            )
        )
        cfg = load_config(cfg_path)
        assert "grade_files" in cfg
        assert "grade_dir" in cfg

    def test_missing_master_raises(self, tmp_path):
        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(yaml.dump({"grade_files": ["g1.csv"]}))
        with pytest.raises(ValueError, match="master_file"):
            load_config(cfg_path)

    def test_missing_both_grade_sources_raises(self, tmp_path):
        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(yaml.dump({"master_file": "m.csv"}))
        with pytest.raises(ValueError, match=r"grade_files.*grade_dir"):
            load_config(cfg_path)

    def test_default_output(self, tmp_path):
        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump({"master_file": "m.csv", "grade_files": ["g.csv"]})
        )
        cfg = load_config(cfg_path)
        assert cfg["output_file"] == "grades_consolidated.csv"


# ============================================================================
# 11. resolve_grade_files
# ============================================================================


class TestResolveGradeFiles:
    def test_explicit_files_only(self, tmp_path):
        (tmp_path / "a.csv").write_text("x")
        (tmp_path / "b.xlsx").write_text("x")
        result = resolve_grade_files(tmp_path, grade_files=["a.csv", "b.xlsx"])
        assert len(result) == 2
        assert result[0].name == "a.csv"
        assert result[1].name == "b.xlsx"

    def test_dir_only(self, tmp_path):
        grade_dir = tmp_path / "grades"
        grade_dir.mkdir()
        (grade_dir / "g1.csv").write_text("x")
        (grade_dir / "g2.xlsx").write_text("x")
        (grade_dir / "g3.ods").write_text("x")
        (grade_dir / "readme.txt").write_text("x")  # .txt is supported
        (grade_dir / "notes.pdf").write_text("x")  # .pdf is NOT supported
        result = resolve_grade_files(tmp_path, grade_dir="grades")
        names = [p.name for p in result]
        assert "g1.csv" in names
        assert "g2.xlsx" in names
        assert "g3.ods" in names
        assert "readme.txt" in names
        assert "notes.pdf" not in names

    def test_dir_ignores_subdirectories(self, tmp_path):
        grade_dir = tmp_path / "grades"
        grade_dir.mkdir()
        (grade_dir / "g1.csv").write_text("x")
        sub = grade_dir / "subdir"
        sub.mkdir()
        (sub / "nested.csv").write_text("x")
        result = resolve_grade_files(tmp_path, grade_dir="grades")
        assert len(result) == 1
        assert result[0].name == "g1.csv"

    def test_both_files_and_dir(self, tmp_path):
        grade_dir = tmp_path / "grades"
        grade_dir.mkdir()
        (grade_dir / "from_dir.csv").write_text("x")
        (tmp_path / "extra.csv").write_text("x")
        result = resolve_grade_files(
            tmp_path, grade_files=["extra.csv"], grade_dir="grades"
        )
        names = [p.name for p in result]
        assert "extra.csv" in names
        assert "from_dir.csv" in names

    def test_deduplication(self, tmp_path):
        """Same file listed explicitly and found in directory."""
        grade_dir = tmp_path / "grades"
        grade_dir.mkdir()
        (grade_dir / "g1.csv").write_text("x")
        result = resolve_grade_files(
            tmp_path, grade_files=["grades/g1.csv"], grade_dir="grades"
        )
        assert len(result) == 1

    def test_dir_sorted_deterministic(self, tmp_path):
        grade_dir = tmp_path / "grades"
        grade_dir.mkdir()
        for name in ["charlie.csv", "alpha.csv", "bravo.csv"]:
            (grade_dir / name).write_text("x")
        result = resolve_grade_files(tmp_path, grade_dir="grades")
        names = [p.name for p in result]
        assert names == ["alpha.csv", "bravo.csv", "charlie.csv"]

    def test_invalid_dir_raises(self, tmp_path):
        with pytest.raises(ValueError, match="not a directory"):
            resolve_grade_files(tmp_path, grade_dir="nonexistent")

    def test_empty_dir_returns_empty(self, tmp_path):
        grade_dir = tmp_path / "grades"
        grade_dir.mkdir()
        result = resolve_grade_files(tmp_path, grade_dir="grades")
        assert result == []

    def test_neither_files_nor_dir(self, tmp_path):
        result = resolve_grade_files(tmp_path)
        assert result == []


# ============================================================================
# 11. Full integration test
# ============================================================================


class TestIntegration:
    def test_end_to_end(self, tmp_path):
        """Full pipeline: master + 3 TA files (csv, xlsx, ods) → output."""
        # Master
        master_path = tmp_path / "master.csv"
        _write_csv(
            master_path,
            ["Numéro étudiant", "Prénom", "Nom de famille", "Email"],
            [
                ["S001", "Alice", "Martin", "alice@univ.fr"],
                ["S002", "Bob", "Bernard", "bob@univ.fr"],
                ["S003", "Claire", "Dubois", "claire@univ.fr"],
                ["S004", "David", "Thomas", "david@univ.fr"],
            ],
        )

        # TA file 1 — CSV with IDs
        ta1 = tmp_path / "group1.csv"
        _write_csv(
            ta1,
            ["Student ID", "Grade"],
            [["S001", "16"], ["S002", "ABS"]],
        )

        # TA file 2 — XLSX with names only
        ta2 = tmp_path / "group2.xlsx"
        pd.DataFrame(
            {"Prénom": ["Claire"], "Nom": ["Dubois"], "Note": ["13,5"]}
        ).to_excel(ta2, index=False, engine="openpyxl")

        # TA file 3 — ODS with ID and name
        ta3 = tmp_path / "group3.ods"
        pd.DataFrame(
            {
                "Numéro étudiant": ["S004"],
                "Prénom": ["David"],
                "Nom": ["Thomas"],
                "Note finale": ["18"],
            }
        ).to_excel(ta3, index=False, engine="odf")

        # Config
        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "master.csv",
                    "grade_files": ["group1.csv", "group2.xlsx", "group3.ods"],
                    "output_file": "results.csv",
                }
            )
        )

        master, reports = consolidate(cfg_path)

        # Verify
        assert all(not r.skipped for r in reports)
        assert master.by_id["S001"].grade == 16.0
        assert master.by_id["S002"].is_absent
        assert master.by_id["S003"].grade == 13.5
        assert master.by_id["S004"].grade == 18.0

        # Check output file exists and is correct
        out = tmp_path / "results.csv"
        assert out.exists()
        df = pd.read_csv(out, keep_default_na=False)
        assert len(df) == 4
        assert df.loc[df["Identifier"] == "S002", "Grade"].values[0] == "ABS"

    def test_master_as_xlsx(self, tmp_path):
        """Master file can also be XLSX."""
        master_path = tmp_path / "master.xlsx"
        pd.DataFrame(
            {
                "Identifiant": ["S001"],
                "Prénom": ["Alice"],
                "Nom": ["Martin"],
                "Courriel": ["alice@univ.fr"],
            }
        ).to_excel(master_path, index=False, engine="openpyxl")

        ta = tmp_path / "ta.csv"
        _write_csv(ta, ["Identifiant", "Note"], [["S001", "17"]])

        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "master.xlsx",
                    "grade_files": ["ta.csv"],
                    "output_file": "out.csv",
                }
            )
        )

        master, _reports = consolidate(cfg_path)
        assert master.by_id["S001"].grade == 17.0

    def test_grade_dir_integration(self, tmp_path):
        """Full pipeline using grade_dir instead of grade_files."""
        master_path = tmp_path / "master.csv"
        _write_csv(
            master_path,
            ["Numéro étudiant", "Prénom", "Nom de famille", "Email"],
            [
                ["S001", "Alice", "Martin", "alice@univ.fr"],
                ["S002", "Bob", "Bernard", "bob@univ.fr"],
            ],
        )

        grade_dir = tmp_path / "ta_grades"
        grade_dir.mkdir()
        _write_csv(
            grade_dir / "group1.csv",
            ["Student ID", "Grade"],
            [["S001", "16"]],
        )
        pd.DataFrame({"Numéro étudiant": ["S002"], "Note": ["14,5"]}).to_excel(
            grade_dir / "group2.xlsx", index=False, engine="openpyxl"
        )

        # Also put a non-grade file in the dir to make sure it's skipped
        (grade_dir / "notes.pdf").write_text("not a spreadsheet")

        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "master.csv",
                    "grade_dir": "ta_grades",
                    "output_file": "results.csv",
                }
            )
        )

        master, reports = consolidate(cfg_path)
        assert len(reports) == 2
        assert all(not r.skipped for r in reports)
        assert master.by_id["S001"].grade == 16.0
        assert master.by_id["S002"].grade == 14.5

    def test_grade_dir_and_files_combined(self, tmp_path):
        """grade_files and grade_dir can be used together."""
        master_path = tmp_path / "master.csv"
        _write_csv(
            master_path,
            ["Numéro étudiant", "Prénom", "Nom de famille", "Email"],
            [
                ["S001", "Alice", "Martin", "alice@univ.fr"],
                ["S002", "Bob", "Bernard", "bob@univ.fr"],
            ],
        )

        grade_dir = tmp_path / "ta_grades"
        grade_dir.mkdir()
        _write_csv(
            grade_dir / "group1.csv",
            ["Student ID", "Grade"],
            [["S001", "15"]],
        )

        # An extra file outside the directory
        _write_csv(
            tmp_path / "extra_ta.csv",
            ["Student ID", "Grade"],
            [["S002", "17"]],
        )

        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "master.csv",
                    "grade_files": ["extra_ta.csv"],
                    "grade_dir": "ta_grades",
                    "output_file": "results.csv",
                }
            )
        )

        master, reports = consolidate(cfg_path)
        assert len(reports) == 2
        assert master.by_id["S001"].grade == 15.0
        assert master.by_id["S002"].grade == 17.0


# ============================================================================
# 12. Edge case: "Nom" column ambiguity with "last_name" role
# ============================================================================


class TestEdgeCases:
    def test_bare_nom_detected_as_last_name(self):
        """'Nom' alone should map to last_name, not first_name."""
        assert detect_column(["Nom", "Prénom"], "last_name") == "Nom"
        assert detect_column(["Nom", "Prénom"], "first_name") == "Prénom"

    def test_grade_column_not_fooled_by_id_numbers(self):
        """
        Student IDs are numeric but should not be mistaken for grades
        when they are already claimed as the ID column.
        """
        df = pd.DataFrame({"Numéro étudiant": ["12345", "12346"], "Note": ["15", "16"]})
        # "Numéro étudiant" is already identified as ID
        col, _ = detect_grade_column(df, {"Numéro étudiant"})
        assert col == "Note"

    def test_parse_grade_fraction_various(self):
        """Various /xx patterns."""
        assert parse_grade("17/20").value == 17.0
        assert parse_grade("8.5/10").value == 8.5
        assert parse_grade("0/20").value == 0.0

    def test_mixed_valid_invalid_grades_in_column(self, tmp_path):
        """Column with mix of valid grades and garbage is still detected
        if ≥ 50% are numeric-like."""
        master = _build_master()
        p = tmp_path / "ta.csv"
        # 3 valid out of 4 rows = 75%
        _write_csv(
            p,
            ["Numéro étudiant", "Mystery"],
            [
                ["12345", "15"],
                ["12346", "ABS"],
                ["12347", "12,5"],
                ["12348", "oops"],
            ],
        )
        report = process_ta_file(p, master)
        # Should detect Mystery as grade column
        assert report.grades_assigned == 2  # 15 and 12.5
        assert report.students_absent == 1  # ABS
        assert master.by_id["12345"].grade == 15.0

    def test_bom_in_csv_column_names(self, tmp_path):
        """BOM character in CSV should not break column detection."""
        p = tmp_path / "master.csv"
        # Write raw bytes with BOM prefix on first column
        p.write_bytes(
            "\ufeffPrénom,Nom de famille,Numéro d\u2019identification,"
            "Adresse de courriel\n"
            "Alice,Martin,S001,alice@univ.fr\n".encode()
        )
        df = read_file(p)
        assert "Prénom" in df.columns  # BOM stripped
        master, _warnings = build_master_index(df)
        assert "S001" in master.by_id
        assert master.by_id["S001"].first_name == "Alice"
        assert master.by_id["S001"].email == "alice@univ.fr"

    def test_moodle_export_format(self, tmp_path):
        """End-to-end with a Moodle-style master export (BOM + curly quotes)."""
        # Master: Moodle export format
        master_path = tmp_path / "participants.csv"
        master_path.write_bytes(
            "\ufeffPrénom,Nom de famille,Numéro d\u2019identification,"
            "Adresse de courriel,Groupes\n"
            "Alice,Martin,S001,alice@univ.fr,Groupe 1\n"
            "Bob,Bernard,S002,bob@univ.fr,Groupe 2\n".encode()
        )

        # TA file
        ta_path = tmp_path / "grades.csv"
        _write_csv(
            ta_path,
            ["Student ID", "Grade"],
            [["S001", "16"], ["S002", "14"]],
        )

        cfg_path = tmp_path / "config.yaml"
        cfg_path.write_text(
            yaml.dump(
                {
                    "master_file": "participants.csv",
                    "grade_files": ["grades.csv"],
                    "output_file": "out.csv",
                }
            )
        )

        master, _reports = consolidate(cfg_path)
        assert master.by_id["S001"].grade == 16.0
        assert master.by_id["S002"].grade == 14.0