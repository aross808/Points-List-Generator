# transform_digital.py
from __future__ import annotations

from typing import Set, Tuple, List
import pandas as pd

from transform_common import (
    truncate_on_first_blank_norm,
    get_col,
    clean_token,
    split_by_io,
    substation_map,
    get_dnp_col,
    dnp_map,
    make_caiso_tags,
)


def transform_digital_kind(
    df: pd.DataFrame,
    *,
    kind: str,                   # "DI" or "DO"
    group_var: str,              # DI="40/2", DO="180/2"
    skip_rows: Set[int] | None = None,
    calculated_desc: Set[str] | None = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:


    skip_rows = skip_rows or set()
    calculated_desc_u = {d.strip().upper() for d in (calculated_desc or set())}

    desc = get_col(df, "pointdescription").astype(str)
    name = clean_token(get_col(df, "pointname"))

    # RAW DNP read + identity mapping (optionally zero out skipped rows)
    raw_dnp = get_dnp_col(df)
    dnp = dnp_map(raw_dnp, skip_rows=set())

    logic_raw = get_col(df, "pointlogic0/1").astype(str)

    parts = logic_raw.str.split("\\\\")
    if parts.map(len).min() >= 2:
        mean0 = parts.str[0].str.strip()
        mean1 = parts.str[1].str.strip()
    else:
        mean0 = logic_raw.str.split("\\\\", n=1).str[0].str.strip()
        mean1 = logic_raw.str.split("\\\\", n=1).str[1].str.strip()

    # CAISO tags keyed by RAW DNP (no compression)
    caiso_prefix = f"CAISO_DNP.{kind}"
    caiso_tags = make_caiso_tags(caiso_prefix, dnp, name)

    # Optional: still blank “calculated” rows (same behavior as analog rewrite)
    if calculated_desc_u:
        for i, d in enumerate(desc.tolist()):
            if str(d).strip().upper() in calculated_desc_u:
                caiso_tags[i] = ""

    # Substation tags: also use RAW DNP (no compression)
    sub_prefix = "BI" if kind == "DI" else "BO"
    rig_tags, scada_tags, sub_idx = substation_map(
        desc,
        name,
        dnp,
        substation_prefix=sub_prefix,
        skip_rows=skip_rows,
        calculated_desc=calculated_desc,
    )

    # -------------------------
    # POINT SELECTION (VISIBLE COLUMNS)
    # -------------------------
    point_selection = pd.DataFrame({
        "Point Description": desc,
        "Point Type": "Digital Input" if kind == "DI" else "Digital Output",
        "Abbreviation": name,
        "RTAC Tag Name": caiso_tags,
        "IP Address": "",
        "0 Means": mean0,
        "1 Means": mean1,
        "Analog Unit in RIG": "",
        "Scaling Applied in SCADA": "",
        "Group/Variation": group_var,
    })

    # meta for index grids (RAW, globally unique)
    point_selection["caiso_kind"] = kind
    point_selection["caiso_index"] = dnp
    point_selection["sub_kind"] = kind
    point_selection["sub_index"] = sub_idx

    # -------------------------
    # CAISO sheet
    # -------------------------
    caiso = pd.DataFrame({
        "Point Description": desc,
        "RIG Tag Name": caiso_tags,
        "RTAC-RIG Source Point(Main)": rig_tags,
        "0 Means": mean0,
        "1 Means": mean1,
        "Group/Variation": group_var,
    })
    caiso["__kind"] = kind
    caiso["__dnp"] = dnp

    # keep only rows that actually got a tag (matches your existing behavior)
    caiso = caiso[caiso["RIG Tag Name"].astype(str).str.strip() != ""]

    # -------------------------
    # Substation sheet
    # -------------------------
    substation = pd.DataFrame({
        "Point Description": desc,
        "RTAC-RIG Tag Name": scada_tags,
        "0 Means": mean0,
        "1 Means": mean1,
        "Group/Variation": group_var,
    })
    substation["__kind"] = kind
    substation["__dnp"] = dnp

    substation = substation[substation["RTAC-RIG Tag Name"].astype(str).str.strip() != ""]

    return point_selection, caiso, substation


def transform_digital_file(
    digital_df: pd.DataFrame,
    *,
    skip_di: Set[int] | None = None,
    skip_do: Set[int] | None = None,
    group_var_di: str = "40/2",
    group_var_do: str = "180/2",
    type_col_norm: str = "pointtype",
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    One digital file contains DI + DO. Split using type_col_norm.
    (Still supported, even though your GUI now uses separate files.)
    """
    di_df, do_df = split_by_io(
        digital_df,
        type_col_norm=type_col_norm,
        input_tokens={"di", "digitalinput", "binaryinput", "input"},
        output_tokens={"do", "digitaloutput", "binaryoutput", "output"},
    )

    ps_frames: List[pd.DataFrame] = []
    c_frames: List[pd.DataFrame] = []
    s_frames: List[pd.DataFrame] = []

    if not di_df.empty:
        ps, c, s = transform_digital_kind(
            di_df,
            kind="DI",
            group_var=group_var_di,
            skip_rows=skip_di,
            calculated_desc=set()
        )
        ps_frames.append(ps); c_frames.append(c); s_frames.append(s)

    if not do_df.empty:
        ps, c, s = transform_digital_kind(
            do_df,
            kind="DO",
            group_var=group_var_do,
            skip_rows=skip_do,
            calculated_desc=set(),
        )
        ps_frames.append(ps); c_frames.append(c); s_frames.append(s)

    return (
        pd.concat(ps_frames, ignore_index=True) if ps_frames else pd.DataFrame(),
        pd.concat(c_frames, ignore_index=True) if c_frames else pd.DataFrame(),
        pd.concat(s_frames, ignore_index=True) if s_frames else pd.DataFrame(),
    )
