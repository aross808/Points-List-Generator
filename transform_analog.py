# transform_analog.py
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


def transform_analog_kind(
    df: pd.DataFrame,
    *,
    kind: str,                   # "AI" or "AO"
    group_var: str,              # AI="32/2", AO="90/1"
    skip_rows: Set[int] | None = None,
    calculated_desc: Set[str] | None = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Returns:
      - point_selection_df (visible columns + meta columns for index grids)
      - caiso_df (with __kind column for sectioned writing)
      - substation_df (with __kind column for sectioned writing)
    """


    print(df.head(8))
    print("Columns:", list(df.columns))



    skip_rows = skip_rows or set()
    calculated_desc_u = {d.strip().upper() for d in (calculated_desc or set())}

    desc = get_col(df, "pointdescription").astype(str)
    name = clean_token(get_col(df, "pointname"))
    unit = clean_token(get_col(df, "eng.units"))

    scale_l = pd.to_numeric(get_col(df, "rige/umaxunits"), errors="coerce")
    scale_h = pd.to_numeric(get_col(df, "rigtocaisomaxcounts"), errors="coerce")
    scaling = (scale_h / scale_l).fillna(0.0)
    scale_fmt = [f"*{v}" for v in scaling]

    # RAW DNP read + identity mapping (optionally zero out skipped rows)
    raw_dnp = get_dnp_col(df)
    dnp = dnp_map(raw_dnp, skip_rows=set())

    # CAISO tags keyed by RAW DNP (no compression)
    caiso_prefix = f"CAISO_DNP.{kind}"
    caiso_tags = make_caiso_tags(caiso_prefix, dnp, name)

    # Substation map (AI/AO) uses RAW index (no compression)
    rig_tags, scada_tags, sub_idx = substation_map(
        desc,
        name,
        dnp,
        substation_prefix=kind,
        skip_rows=skip_rows,
        calculated_desc=calculated_desc,
    )

    # -------------------------
    # POINT SELECTION (VISIBLE COLUMNS)
    # -------------------------
    point_selection = pd.DataFrame({
        "Point Description": desc,
        "Point Type": "Analog Input" if kind == "AI" else "Analog Output",
        "Abbreviation": name,
        "RTAC Tag Name": caiso_tags,
        "IP Address": "",
        "0 Means": "",
        "1 Means": "",
        "Analog Unit in RIG": unit,
        "Scaling Applied in SCADA": scale_fmt,
        "Group/Variation": group_var,
    })

    # meta for index grids (RAW, globally unique)
    point_selection["caiso_kind"] = kind
    point_selection["caiso_index"] = dnp
    point_selection["sub_kind"] = kind
    point_selection["sub_index"] = sub_idx  # "" for skipped/calculated/invalid rows

    # -------------------------
    # CAISO sheet
    # -------------------------
    caiso = pd.DataFrame({
        "Point Description": desc,
        "RIG Tag Name": caiso_tags,
        "RTAC-RIG Source Point(Main)": rig_tags,
        "Analog Unit to CAISO": unit,
        "Scaling Applied in RIG": "*1",
        "Group/Variation": group_var,
    })
    caiso["__kind"] = kind
    caiso["__dnp"] = dnp

    # keep only rows that actually got a tag
    caiso = caiso[caiso["RIG Tag Name"].astype(str).str.strip() != ""]

    # -------------------------
    # Substation sheet
    # -------------------------
    substation = pd.DataFrame({
        "Point Description": desc,
        "RTAC-RIG Tag Name": scada_tags,
        "Analog Unit in SUBSTATION RTU": unit,
        "Scaling Applied in SUBSTATION RTU": scale_fmt,
        "Group/Variation": group_var,
    })
    substation["__kind"] = kind
    substation["__dnp"] = dnp

    substation = substation[substation["RTAC-RIG Tag Name"].astype(str).str.strip() != ""]

    return point_selection, caiso, substation


def transform_analog_file(
    analog_df: pd.DataFrame,
    *,
    skip_ai: Set[int] | None = None,
    skip_ao: Set[int] | None = None,
    group_var_ai: str = "32/2",
    group_var_ao: str = "90/1",
    type_col_norm: str = "pointtype",
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    One analog file contains AI + AO. Split using type_col_norm.
    (Still supported, even though your GUI now uses separate files.)
    """
    ai_df, ao_df = split_by_io(
        analog_df,
        type_col_norm=type_col_norm,
        input_tokens={"ai", "analoginput", "input"},
        output_tokens={"ao", "analogoutput", "output"},
    )

    ps_frames: List[pd.DataFrame] = []
    c_frames: List[pd.DataFrame] = []
    s_frames: List[pd.DataFrame] = []

    if not ai_df.empty:
        ps, c, s = transform_analog_kind(
            ai_df,
            kind="AI",
            group_var=group_var_ai,
            skip_rows=skip_ai,
            calculated_desc=set(),
        )
        ps_frames.append(ps); c_frames.append(c); s_frames.append(s)

    if not ao_df.empty:
        ps, c, s = transform_analog_kind(
            ao_df,
            kind="AO",
            group_var=group_var_ao,
            skip_rows=skip_ao,
            calculated_desc=set()
        )
        ps_frames.append(ps); c_frames.append(c); s_frames.append(s)

    return (
        pd.concat(ps_frames, ignore_index=True) if ps_frames else pd.DataFrame(),
        pd.concat(c_frames, ignore_index=True) if c_frames else pd.DataFrame(),
        pd.concat(s_frames, ignore_index=True) if s_frames else pd.DataFrame(),
    )
