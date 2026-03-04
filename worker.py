# worker.py
from __future__ import annotations

import traceback
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any

from PyQt6 import QtCore

from io_reader import read_with_detected_header, InputReadError
from transform_common import parse_skip_list
from transform_analog import transform_analog_kind
from transform_digital import transform_digital_kind
from excel_writer import write_workbook


@dataclass(frozen=True)
class JobConfig:
    # Separate files (no PointType column needed)
    ai_path: Optional[str]
    ao_path: Optional[str]
    di_path: Optional[str]
    do_path: Optional[str]
    output_path: str

    # GUI text fields like: "1,2,4,9" or "1-3,8"
    skip_ai_text: str = ""
    skip_ao_text: str = ""
    skip_di_text: str = ""
    skip_do_text: str = ""

    meters: List[Dict[str, Any]] = field(default_factory=list)


class GenerateWorker(QtCore.QThread):
    log_msg = QtCore.pyqtSignal(str)
    finished_ok = QtCore.pyqtSignal(str)
    finished_err = QtCore.pyqtSignal(str)

    def __init__(self, cfg: JobConfig):
        super().__init__()
        self.cfg = cfg

    def _log(self, msg: str) -> None:
        self.log_msg.emit(msg)

    def run(self) -> None:
        try:
            self._log("Starting…")

            # -------------------------
            # Read inputs (Step 2)
            # -------------------------
            self._log("Reading files…")

            # Analog requirements (transform uses eng.units + scaling columns)
            analog_required = [
                "pointdescription",
                "pointname",
            ]

            # Digital requirements (transform uses pointlogic0/1 too)
            digital_required = [
                "pointdescription",
                "pointname",
                "dnpindex(0notused)",
                "pointlogic0/1",
            ]

            ai_df = read_with_detected_header(
                self.cfg.ai_path,
                required_headers=analog_required,
                scan_rows=25,
            ) if self.cfg.ai_path else None

            ao_df = read_with_detected_header(
                self.cfg.ao_path,
                required_headers=analog_required,
                scan_rows=25,
            ) if self.cfg.ao_path else None

            di_df = read_with_detected_header(
                self.cfg.di_path,
                required_headers=digital_required,
                scan_rows=25,
            ) if self.cfg.di_path else None

            do_df = read_with_detected_header(
                self.cfg.do_path,
                required_headers=digital_required,
                scan_rows=25,
            ) if self.cfg.do_path else None

            if (
                (ai_df is None or ai_df.empty)
                and (ao_df is None or ao_df.empty)
                and (di_df is None or di_df.empty)
                and (do_df is None or do_df.empty)
            ):
                self.finished_err.emit("No input files selected (AI/AO/DI/DO are all empty).")
                return

            # -------------------------
            # Parse skip lists (GUI -> sets)
            # -------------------------
            skip_ai = parse_skip_list(self.cfg.skip_ai_text)
            skip_ao = parse_skip_list(self.cfg.skip_ao_text)
            skip_di = parse_skip_list(self.cfg.skip_di_text)
            skip_do = parse_skip_list(self.cfg.skip_do_text)

            # -------------------------
            # Transform (Step 4)
            # -------------------------
            self._log("Processing data...")

            point_frames = []
            caiso_frames = []
            sub_frames = []

            if ai_df is not None and not ai_df.empty:
                ps, c, s = transform_analog_kind(
                    ai_df,
                    kind="AI",
                    group_var="32/2",
                    skip_rows=skip_ai,
                    calculated_desc={"RIG HEARTBEAT COUNTER"},
                )
                if not ps.empty: point_frames.append(ps)
                if not c.empty: caiso_frames.append(c)
                if not s.empty: sub_frames.append(s)
                self._log("Analog Inputs Written")

            if ao_df is not None and not ao_df.empty:
                ps, c, s = transform_analog_kind(
                    ao_df,
                    kind="AO",
                    group_var="90/1",
                    skip_rows=skip_ao,
                    calculated_desc=set(),
                )
                if not ps.empty: point_frames.append(ps)
                if not c.empty: caiso_frames.append(c)
                if not s.empty: sub_frames.append(s)
                self._log("Analog Outputs Written")

            if di_df is not None and not di_df.empty:
                ps, c, s = transform_digital_kind(
                    di_df,
                    kind="DI",
                    group_var="40/2",
                    skip_rows=skip_di,
                    calculated_desc={"AGG UNIT CONNECTION STATUS"},
                )
                if not ps.empty: point_frames.append(ps)
                if not c.empty: caiso_frames.append(c)
                if not s.empty: sub_frames.append(s)
                self._log("Digital Inputs Written")

            if do_df is not None and not do_df.empty:
                ps, c, s = transform_digital_kind(
                    do_df,
                    kind="DO",
                    group_var="180/2",
                    skip_rows=skip_do,
                    calculated_desc=set(),
                )
                if not ps.empty: point_frames.append(ps)
                if not c.empty: caiso_frames.append(c)
                if not s.empty: sub_frames.append(s)
                self._log("Digital Outputs Written")

            if not point_frames and not caiso_frames and not sub_frames:
                self.finished_err.emit("No data processed (empty outputs). Check input files.")
                return

            import pandas as pd
            point_df = pd.concat(point_frames, ignore_index=True) if point_frames else None
            caiso_df = pd.concat(caiso_frames, ignore_index=True) if caiso_frames else None
            sub_df = pd.concat(sub_frames, ignore_index=True) if sub_frames else None

            # -------------------------
            # Write workbook (Step 5)
            # -------------------------
            self._log("Writing...")
            self._log(f"Meters Written: {self.cfg.meters}")
            out = write_workbook(
                self.cfg.output_path,
                point_selection_df=point_df,
                caiso_df=caiso_df,
                substation_df=sub_df,
                meters=self.cfg.meters,
                autosize=True,
            )
            self._log("SUCCESS")
            self.finished_ok.emit(out)

        except InputReadError as e:
            self.finished_err.emit(str(e))

        except Exception as e:
            tb = traceback.format_exc()
            self.finished_err.emit(f"Exception: {e}\n\n{tb}")
