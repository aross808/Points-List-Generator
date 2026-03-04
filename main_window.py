# main_window.py
from __future__ import annotations

import os
from typing import Optional

from transform_common import parse_skip_list

try:
    from PyQt6 import QtCore, QtWidgets
    using_pyqt6 = True
except Exception:
    from PyQt5 import QtCore, QtWidgets  # type: ignore
    using_pyqt6 = False

from worker import JobConfig, GenerateWorker


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ulteig Point List Generator — GUI")
        self.resize(980, 680)

        # Separate files (no PointType column)
        self.ai_path: Optional[str] = None
        self.ao_path: Optional[str] = None
        self.di_path: Optional[str] = None
        self.do_path: Optional[str] = None

        self.output_path: str = os.path.abspath("output_combined.xlsx")
        self.worker: Optional[GenerateWorker] = None

        self._build_ui()

    # -------------------------
    # UI
    # -------------------------
    def _build_ui(self):
        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)

        # Title
        title = QtWidgets.QLabel("CAISO Points List Generator")
        font = title.font()
        font.setPointSize(18)
        font.setBold(True)
        title.setFont(font)
        title.setAlignment(
            QtCore.Qt.AlignmentFlag.AlignHCenter if using_pyqt6 else QtCore.Qt.AlignHCenter
        )
        layout.addWidget(title)

        # -------------------------
        # File pickers (AI/AO/DI/DO)
        # -------------------------
        grid = QtWidgets.QGridLayout()
        layout.addLayout(grid)

        row = 0

        grid.addWidget(QtWidgets.QLabel("AI File (.xlsx/.csv):"), row, 0)
        self.txtAI = QtWidgets.QLineEdit()
        self.txtAI.setPlaceholderText("Select AI file (Analog Inputs)")
        grid.addWidget(self.txtAI, row, 1)
        btnAI = QtWidgets.QPushButton("Browse…")
        btnAI.clicked.connect(self.pick_ai)
        grid.addWidget(btnAI, row, 2)
        row += 1

        grid.addWidget(QtWidgets.QLabel("AO File (.xlsx/.csv):"), row, 0)
        self.txtAO = QtWidgets.QLineEdit()
        self.txtAO.setPlaceholderText("Select AO file (Analog Outputs)")
        grid.addWidget(self.txtAO, row, 1)
        btnAO = QtWidgets.QPushButton("Browse…")
        btnAO.clicked.connect(self.pick_ao)
        grid.addWidget(btnAO, row, 2)
        row += 1

        grid.addWidget(QtWidgets.QLabel("DI File (.xlsx/.csv):"), row, 0)
        self.txtDI = QtWidgets.QLineEdit()
        self.txtDI.setPlaceholderText("Select DI file (Digital Inputs)")
        grid.addWidget(self.txtDI, row, 1)
        btnDI = QtWidgets.QPushButton("Browse…")
        btnDI.clicked.connect(self.pick_di)
        grid.addWidget(btnDI, row, 2)
        row += 1

        grid.addWidget(QtWidgets.QLabel("DO File (.xlsx/.csv):"), row, 0)
        self.txtDO = QtWidgets.QLineEdit()
        self.txtDO.setPlaceholderText("Select DO file (Digital Outputs)")
        grid.addWidget(self.txtDO, row, 1)
        btnDO = QtWidgets.QPushButton("Browse…")
        btnDO.clicked.connect(self.pick_do)
        grid.addWidget(btnDO, row, 2)
        row += 1

        grid.addWidget(QtWidgets.QLabel("Output Excel:"), row, 0)
        self.txtOutput = QtWidgets.QLineEdit()
        self.txtOutput.setText(self.output_path)
        grid.addWidget(self.txtOutput, row, 1)
        btnOut = QtWidgets.QPushButton("Choose…")
        btnOut.clicked.connect(self.pick_output)
        grid.addWidget(btnOut, row, 2)
        row += 1

        # -------------------------
        # Meters (DNP-based)
        # -------------------------
        meters_box = QtWidgets.QGroupBox("Meters (defined by CAISO DNP indices)")
        meters_layout = QtWidgets.QVBoxLayout(meters_box)

        self.meter_container = QtWidgets.QVBoxLayout()
        meters_layout.addLayout(self.meter_container)

        btn_add_meter = QtWidgets.QPushButton("+ Add Meter")
        meters_layout.addWidget(btn_add_meter)

        def add_meter_widget():
            meter_box = QtWidgets.QGroupBox("Meter")
            g = QtWidgets.QGridLayout(meter_box)

            g.addWidget(QtWidgets.QLabel("Name:"), 0, 0)
            txtName = QtWidgets.QLineEdit()
            txtName.setPlaceholderText("e.g. MAIN F1")
            g.addWidget(txtName, 0, 1)

            g.addWidget(QtWidgets.QLabel("DNP Indices:"), 1, 0)
            txtDNP = QtWidgets.QLineEdit()
            txtDNP.setPlaceholderText("e.g. 1-3,7,12")
            g.addWidget(txtDNP, 1, 1)

            meter_box._fields = {"name": txtName, "dnp": txtDNP}
            self.meter_container.addWidget(meter_box)

        btn_add_meter.clicked.connect(add_meter_widget)
        layout.addWidget(meters_box)

        # -------------------------
        # Run row
        # -------------------------
        btn_row = QtWidgets.QHBoxLayout()
        layout.addLayout(btn_row)

        self.btnRun = QtWidgets.QPushButton("Generate Excel")
        self.btnRun.clicked.connect(self.on_run)
        btn_row.addWidget(self.btnRun)

        self.lblStatus = QtWidgets.QLabel("")
        btn_row.addWidget(self.lblStatus)
        btn_row.addStretch(1)

        # Log
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, stretch=1)

    # -------------------------
    # Pickers
    # -------------------------
    def pick_ai(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select AI file",
            "",
            "Excel/CSV (*.xlsx *.xlsm *.xls *.csv);;All Files (*)",
        )
        if path:
            self.ai_path = path
            self.txtAI.setText(path)

    def pick_ao(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select AO file",
            "",
            "Excel/CSV (*.xlsx *.xlsm *.xls *.csv);;All Files (*)",
        )
        if path:
            self.ao_path = path
            self.txtAO.setText(path)

    def pick_di(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select DI file",
            "",
            "Excel/CSV (*.xlsx *.xlsm *.xls *.csv);;All Files (*)",
        )
        if path:
            self.di_path = path
            self.txtDI.setText(path)

    def pick_do(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select DO file",
            "",
            "Excel/CSV (*.xlsx *.xlsm *.xls *.csv);;All Files (*)",
        )
        if path:
            self.do_path = path
            self.txtDO.setText(path)

    def pick_output(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Choose output Excel file",
            self.txtOutput.text().strip() or "output_combined.xlsx",
            "Excel (*.xlsx)",
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_path = path
            self.txtOutput.setText(path)

    # -------------------------
    # Worker wiring
    # -------------------------
    def append_log(self, msg: str):
        self.log.appendPlainText(msg)

    def collect_meters(self):
        """
        Returns a list of meters from the UI.
        Each meter: {"label": str, "dnp": Set[int]}
        Labels can repeat. DNP indices can overlap across meters.
        """
        meters = []
        for i in range(self.meter_container.count()):
            w = self.meter_container.itemAt(i).widget()
            if w is None or not hasattr(w, "_fields"):
                continue

            fields = w._fields
            label = fields["name"].text().strip()
            dnp_text = fields["dnp"].text().strip()

            if not label or not dnp_text:
                continue

            dnp_set = parse_skip_list(dnp_text)  # supports: "1-3,7,12"
            if not dnp_set:
                continue

            meters.append({"label": label, "dnp": dnp_set})

        return meters

    def on_run(self):
        # Pull paths from UI
        self.ai_path = self.txtAI.text().strip() or None
        self.ao_path = self.txtAO.text().strip() or None
        self.di_path = self.txtDI.text().strip() or None
        self.do_path = self.txtDO.text().strip() or None
        self.output_path = self.txtOutput.text().strip() or os.path.abspath("output_combined.xlsx")

        m = self.collect_meters()
        self.append_log(f"METERS: {m}")
        self.append_log("")

        cfg = JobConfig(
            ai_path=self.ai_path,
            ao_path=self.ao_path,
            di_path=self.di_path,
            do_path=self.do_path,
            output_path=self.output_path,
            meters=m,
            skip_ai_text="",
            skip_ao_text="",
            skip_di_text="",
            skip_do_text="",
        )

        self.btnRun.setEnabled(False)
        self.lblStatus.setText("Running…")

        self.worker = GenerateWorker(cfg)
        self.worker.log_msg.connect(self.append_log)
        self.worker.finished_ok.connect(self.on_done)
        self.worker.finished_err.connect(self.on_err)
        self.worker.start()

    def on_done(self, out_path: str):
        self.btnRun.setEnabled(True)
        self.lblStatus.setText(f"Done: {out_path}")
        self.append_log(f"\nDONE: {out_path}")

    def on_err(self, err: str):
        self.btnRun.setEnabled(True)
        self.lblStatus.setText("Error")
        self.append_log(f"\nERROR:\n{err}")


def run_app():
    import sys
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec() if using_pyqt6 else app.exec_())
