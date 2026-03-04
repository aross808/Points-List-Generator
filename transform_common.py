# transform_common.py
from __future__ import annotations

from typing import Optional, Set, Tuple, List, Iterable
import pandas as pd

from io_reader import col_index_map, _norm


# ----------------------------
# Step 3: truncate on blank description
# ----------------------------
def truncate_on_first_blank_norm(df: Optional[pd.DataFrame], norm_col: str) -> Optional[pd.DataFrame]:
    """
    New behavior:
      1) drop leading blank rows in norm_col
      2) once data starts, stop at the first blank row
    """
    if df is None or df.empty:
        return df

    cols = col_index_map(df)
    # NOTE: norm_col should be provided in normalized form already (e.g. "pointdescription")
    if norm_col not in cols:
        return df

    s = df.iloc[:, cols[norm_col]]
    blank = s.isna() | (s.astype(str).str.strip() == "")
    if not blank.any():
        return df

    blank_arr = blank.to_numpy()

    # 1) find first non-blank (start of real data)
    nonblank_positions = (~blank_arr).nonzero()[0]
    if len(nonblank_positions) == 0:
        return df.iloc[:0]  # all blank

    start = int(nonblank_positions[0])

    # 2) find first blank AFTER start
    after = blank_arr[start:]
    if after.any():
        stop = start + int(after.argmax())  # first True within slice
        return df.iloc[start:stop]

    return df.iloc[start:]


def get_col(df: pd.DataFrame, norm_col: str, aliases: Iterable[str] = ()) -> pd.Series:
    """
    Column getter using the SAME normalization as io_reader (_norm).
    This makes lookups tolerant to punctuation/spacing differences in Excel headers.
    """
    cols = col_index_map(df)

    key = _norm(norm_col)
    if key in cols:
        return df.iloc[:, cols[key]]

    for a in aliases:
        ak = _norm(a)
        if ak in cols:
            return df.iloc[:, cols[ak]]

    raise ValueError(
        f"Missing required column: {norm_col}. Available: {sorted(cols.keys())}"
    )


def get_dnp_col(df: pd.DataFrame) -> pd.Series:
    """
    Robust DNP Index column getter.
    Accepts common header variants that normalize differently across exports.
    """
    return get_col(
        df,
        "dnpindex",
        aliases=(
            "dnpindex(0notused)",
            "dnpindex0notused",
            "dnp index",
            "dnp",
        ),
    )


def clean_token(s: pd.Series) -> pd.Series:
    return s.astype(str).str.replace(" ", "_", regex=False)


# ----------------------------
# GUI skip list parsing: "1,2,4,9" or "1-3, 8"
# ----------------------------
def parse_skip_list(text: str) -> Set[int]:
    s = (text or "").strip()
    if not s:
        return set()

    out: Set[int] = set()
    parts = [p.strip() for p in s.split(",") if p.strip()]

    for p in parts:
        if "-" in p:
            a, b = [x.strip() for x in p.split("-", 1)]
            if a.isdigit() and b.isdigit():
                lo, hi = int(a), int(b)
                if lo > hi:
                    lo, hi = hi, lo
                lo = max(lo, 1)
                if hi >= 1:
                    out.update(range(lo, hi + 1))
        else:
            if p.isdigit():
                n = int(p)
                if n >= 1:
                    out.add(n)
    return out


# ----------------------------
# Split one file into input/output based on a type column
# ----------------------------
def split_by_io(
    df: pd.DataFrame,
    *,
    type_col_norm: str,
    input_tokens: Set[str],
    output_tokens: Set[str],
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns (inputs_df, outputs_df) based on a column in the sheet.

    If the column doesn't exist, this treats ALL rows as "input" (safe default).
    """
    cols = col_index_map(df)
    if type_col_norm not in cols:
        return df.copy(), df.iloc[0:0].copy()

    t = df.iloc[:, cols[type_col_norm]].astype(str).str.strip().str.lower()

    def tok(x: str) -> str:
        return x.replace(" ", "")

    t_norm = t.map(tok)
    in_set = {tok(x) for x in input_tokens}
    out_set = {tok(x) for x in output_tokens}

    in_mask = t_norm.isin(in_set)
    out_mask = t_norm.isin(out_set)

    # Anything unrecognized defaults to input
    inputs_df = df[in_mask | (~out_mask)].copy()
    outputs_df = df[out_mask].copy()
    return inputs_df, outputs_df


# ----------------------------
# DNP maps (NO compression, NO legacy overwrite)
# ----------------------------
def dnp_map(dnp: pd.Series, *, skip_rows: Set[int] | None = None) -> pd.Series:
    """
    Identity map for CAISO DNP indices (NO compression).

    - Coerces non-numeric/blank to 0
    - Optionally sets skipped rows to 0 (skip_rows are 1-based after truncation)
    """
    out = pd.to_numeric(dnp, errors="coerce").fillna(0).astype(int)

    skip_rows = skip_rows or set()
    if skip_rows:
        # build a boolean mask in positional space (safe for any index)
        mask = [False] * len(out)
        for r in skip_rows:
            i = r - 1
            if 0 <= i < len(mask):
                mask[i] = True
        out = out.mask(pd.Series(mask, index=out.index), 0)

    return out


def make_caiso_tags(prefix: str, dnp, abbr) -> List[str]:
    """
    Build CAISO tag names keyed by RAW DNP index (NO compression).
    Accepts Series or list-like. Blank tag if index <= 0.

    This version is hardened to avoid dtype/to-list weirdness.
    """
    dnp_s = pd.Series(dnp)
    abbr_s = pd.Series(abbr)

    dnp_i = pd.to_numeric(dnp_s, errors="coerce").fillna(0).astype(int)
    abbr_t = abbr_s.astype(str)

    tags: List[str] = []
    for ii, a in zip(dnp_i.tolist(), abbr_t.tolist()):
        ii = int(ii)
        tags.append(f"{prefix}_{ii:04d}_{a}" if ii > 0 else "")
    return tags


# ----------------------------
# Substation tag/index builder (NO COMPRESSION)
# ----------------------------
def substation_map(
    desc: pd.Series,
    abbr: pd.Series,
    dnp: pd.Series,
    *,
    substation_prefix: str,              # e.g. "AI","AO","BI","BO"
    skip_rows: Set[int] | None = None,   # 1-based row numbers (post-truncation)
    calculated_desc: Set[str] | None = None,
) -> Tuple[List[str], List[str], List[object]]:
    """
    Builds (rig_tags, scada_tags, sub_idx) aligned to input rows.

    IMPORTANT:
      - Uses the RAW DNP index from the input file (no per-kind renumbering/compression).
      - If a row is skipped or calculated, outputs "", "", "".
      - If DNP is 0/blank/non-numeric, outputs "", "", "".
    """
    skip_rows = skip_rows or set()
    calculated_desc = {d.strip().upper() for d in (calculated_desc or set())}

    rig_tags: List[str] = []
    scada_tags: List[str] = []
    sub_idx: List[object] = []

    dnp_i = pd.to_numeric(dnp, errors="coerce").fillna(0).astype(int)

    for row_num, (d, a, raw_idx) in enumerate(zip(desc, abbr, dnp_i), start=1):
        du = str(d).strip().upper()

        if row_num in skip_rows:
            rig_tags.append(""); scada_tags.append(""); sub_idx.append("")
            continue

        if du in calculated_desc:
            rig_tags.append("CALCULATED BY THE RIG"); scada_tags.append(""); sub_idx.append("")
            continue

        if int(raw_idx) <= 0:
            rig_tags.append(""); scada_tags.append(""); sub_idx.append("")
            continue

        tag = f"SUBSTATION_RTU_DNP.{substation_prefix}_{int(raw_idx):04d}_{a}"
        rig_tags.append(tag)
        scada_tags.append(tag)
        sub_idx.append(int(raw_idx))

    return rig_tags, scada_tags, sub_idx
