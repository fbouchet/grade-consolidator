"""
consolidate_grades
==================
Consolidate student grades from multiple TA files (CSV / XLSX / ODS)
into a single Moodle-compatible CSV.
"""

from consolidate_grades.consolidate import (
    ABSENT_TOKENS,
    COLUMN_ALIASES,
    SUPPORTED_EXTENSIONS,
    FileReport,
    MasterIndex,
    ParsedGrade,
    Student,
    build_master_index,
    clean_column_names,
    consolidate,
    detect_column,
    detect_grade_column,
    load_config,
    main,
    make_name_key,
    normalize_name,
    normalize_text,
    parse_grade,
    process_ta_file,
    read_file,
    resolve_grade_files,
    write_moodle_csv,
)

__all__ = [
    "ABSENT_TOKENS",
    "COLUMN_ALIASES",
    "SUPPORTED_EXTENSIONS",
    "FileReport",
    "MasterIndex",
    "ParsedGrade",
    "Student",
    "build_master_index",
    "clean_column_names",
    "consolidate",
    "detect_column",
    "detect_grade_column",
    "load_config",
    "main",
    "make_name_key",
    "normalize_name",
    "normalize_text",
    "parse_grade",
    "process_ta_file",
    "read_file",
    "resolve_grade_files",
    "write_moodle_csv",
]