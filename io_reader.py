# io_reader.py
from __future__ import annotations

import os
from typing import Optional, Sequence

import re 

import pandas as pd


class InputReadError(RuntimeError):
    """Raised when a file exists but cannot be read/parsed as expected."""
    pass


_norm_re = re.compile(r"[^a-z0-9]+")

def _norm(x) -> str:
    """
    Aggressive header normalization:
      - lowercase + strip
      - convert newlines/tabs to spaces
      - remove ALL non-alphanumeric chars (spaces, punctuation, slashes, dots, parentheses)
    Examples:
      "ENG. Units"            -> "engunits"
      "eng.units"             -> "engunits"
      "DNP Index (0 Not Used)"-> "dnpindex0notused"
      "Point Logic 0/1"       -> "pointlogic01"
    """
    if x is None:
        return ""
    s = str(x).strip().lower()
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = _norm_re.sub("", s)
    return s



def _resolve_path(path: str) -> str:
    """Clean quotes/whitespace; if no extension, assume .xlsx; return absolute path."""
    p = path.strip().strip('"').strip("'")
    if not os.path.splitext(p)[1]:
        p += ".xlsx"
    return os.path.abspath(p)


def _read_raw_no_header(path: str) -> pd.DataFrame:
    """Read file with header=None so we can scan rows for the real header row."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, header=None)
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, header=None)
    raise InputReadError(f"Unsupported file type: {ext} (expected .csv/.xlsx/.xlsm/.xls)")


def find_header_row(raw_df: pd.DataFrame, required_headers: Sequence[str], scan_rows: int = 20) -> int:
    """
    Scan the first `scan_rows` rows of a raw DataFrame and return the row index
    that contains all required headers (after normalization).
    """
    required = [_norm(h) for h in required_headers]

    limit = min(scan_rows, len(raw_df))
    for r in range(limit):
        row_vals = raw_df.iloc[r].tolist()
        row_norm = [_norm(v) for v in row_vals]
        if all(req in row_norm for req in required):
            return r

    raise InputReadError(
        f"Header row not found in first {limit} rows. Required headers: {list(required_headers)}"
    )


def read_file(path: Optional[str]) -> Optional[pd.DataFrame]:
    """
    Minimal reader: reads CSV/Excel using default header behavior.
    Returns None if path is empty or file missing.
    """
    if not path or not str(path).strip():
        return None

    p = _resolve_path(path)
    if not os.path.exists(p):
        return None

    ext = os.path.splitext(p)[1].lower()
    if ext == ".csv":
        return pd.read_csv(p)
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(p)
    raise InputReadError(f"Unsupported file type: {ext} (expected .csv/.xlsx/.xlsm/.xls)")


def read_with_detected_header(
    path: Optional[str],
    required_headers: Sequence[str],
    scan_rows: int = 20,
) -> Optional[pd.DataFrame]:
    """
    Reads CSV/Excel and auto-detects the header row by scanning the first `scan_rows`.
    Returns None if path is empty or file missing.
    Raises InputReadError if file exists but a valid header row cannot be found.
    """
    if not path or not str(path).strip():
        return None

    p = _resolve_path(path)
    if not os.path.exists(p):
        return None

    raw = _read_raw_no_header(p)
    header_row = find_header_row(raw, required_headers, scan_rows)

    ext = os.path.splitext(p)[1].lower()
    if ext == ".csv":
        return pd.read_csv(p, header=header_row)
    return pd.read_excel(p, header=header_row)


def col_index_map(df: pd.DataFrame) -> dict[str, int]:
    """
    Convenience lookup: normalized_col_name -> positional index in the dataframe.
    Example:
        cols = col_index_map(df)
        i_desc = cols["pointdescription"]
    """
    return {_norm(c): i for i, c in enumerate(df.columns)}
