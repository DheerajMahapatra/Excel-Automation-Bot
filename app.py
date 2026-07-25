# # # # # """
# # # # # Excel Auto-Fill Bot - Desktop App

# # # # # Pick (or drag & drop) an Excel file (multi-sheet, dynamic headers - nothing
# # # # # hardcoded) and a JSON or XML data file (e.g. an Aadhaar-style API response).
# # # # # The app fuzzy-matches the data's keys against each sheet's headers, fills
# # # # # matching values into a brand new row of every sheet, writes "N/A" for
# # # # # anything with no matching data, and overwrites the same Excel file.

# # # # # - Har record (chahe duplicate hi kyu na ho) apni alag row leta hai - kuch bhi
# # # # #   skip/merge nahi hota.
# # # # # - By default koi bhi sheet skip nahi hoti - jis sheet mein match nahi milta
# # # # #   wahan bhi ek row jaake N/A se bhar jaati hai. "Sirf unrelated sheets skip
# # # # #   karo" checkbox se old (strict) behaviour wapas mil jata hai.
# # # # # """

# # # # # import os
# # # # # import sys
# # # # # import traceback

# # # # # from PySide6.QtCore import Qt, QThread, Signal
# # # # # from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent, QColor, QIcon, QPixmap
# # # # # from PySide6.QtWidgets import (
# # # # #     QApplication,
# # # # #     QCheckBox,
# # # # #     QFileDialog,
# # # # #     QFrame,
# # # # #     QGraphicsDropShadowEffect,
# # # # #     QHBoxLayout,
# # # # #     QLabel,
# # # # #     QMainWindow,
# # # # #     QPushButton,
# # # # #     QSizePolicy,
# # # # #     QTextEdit,
# # # # #     QVBoxLayout,
# # # # #     QWidget,
# # # # # )

# # # # # from filler_core import fill_workbook

# # # # # # ---------------------------------------------------------------------------
# # # # # # Palette (matches the Excel Auto-Fill Bot design spec)
# # # # # # ---------------------------------------------------------------------------
# # # # # BG = "#0b1326"
# # # # # SURFACE = "#0b1326"
# # # # # PANEL = "#171f33"
# # # # # PANEL_LIGHT = "#1e2338"
# # # # # SURFACE_HIGH = "#222a3d"
# # # # # SURFACE_HIGHEST = "#2d3449"
# # # # # BORDER = "#464554"
# # # # # BORDER_ACTIVE = "#8083ff"
# # # # # ACCENT = "#c0c1ff"                 # primary
# # # # # ACCENT_CONTAINER = "#8083ff"       # primary-container
# # # # # ACCENT_HOVER = "#6f72e8"
# # # # # ACCENT_SOFT = "#2a2d55"
# # # # # SECONDARY = "#89ceff"
# # # # # SECONDARY_CONTAINER = "#00a2e6"
# # # # # TERTIARY = "#4edea3"
# # # # # TERTIARY_CONTAINER = "#00885d"
# # # # # TEXT_PRIMARY = "#dae2fd"
# # # # # TEXT_SECONDARY = "#c7c4d7"
# # # # # TEXT_MUTED = "#908fa0"
# # # # # SUCCESS = "#4edea3"
# # # # # ERROR = "#ffb4ab"

# # # # # ASSET_DIR = os.path.dirname(os.path.abspath(__file__))
# # # # # LOGO_PATH = os.path.join(ASSET_DIR, "logo.png")


# # # # # class DropZone(QFrame):
# # # # #     """Clickable + drag-and-drop file picker box (bento-grid card)."""

# # # # #     def __init__(self, icon, icon_color, title, hint, extensions, badge_text, on_pick):
# # # # #         super().__init__()
# # # # #         self.extensions = extensions
# # # # #         self.on_pick = on_pick
# # # # #         self.path = None

# # # # #         self.setAcceptDrops(True)
# # # # #         self.setCursor(Qt.PointingHandCursor)
# # # # #         self.setMinimumHeight(170)
# # # # #         self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
# # # # #         self._set_style(active=False)

# # # # #         shadow = QGraphicsDropShadowEffect(self)
# # # # #         shadow.setBlurRadius(28)
# # # # #         shadow.setOffset(0, 8)
# # # # #         shadow.setColor(QColor(0, 0, 0, 110))
# # # # #         self.setGraphicsEffect(shadow)

# # # # #         layout = QVBoxLayout(self)
# # # # #         layout.setContentsMargins(20, 24, 20, 24)
# # # # #         layout.setSpacing(8)
# # # # #         layout.setAlignment(Qt.AlignCenter)

# # # # #         icon_wrap = QLabel(icon)
# # # # #         icon_wrap.setAlignment(Qt.AlignCenter)
# # # # #         icon_wrap.setFixedSize(64, 64)
# # # # #         icon_wrap.setStyleSheet(
# # # # #             f"font-size: 26px; background-color: {SURFACE_HIGH}; border-radius: 18px; color: {icon_color};"
# # # # #         )

# # # # #         icon_row = QHBoxLayout()
# # # # #         icon_row.addStretch()
# # # # #         icon_row.addWidget(icon_wrap)
# # # # #         icon_row.addStretch()

# # # # #         self.title_label = QLabel(title)
# # # # #         self.title_label.setAlignment(Qt.AlignCenter)
# # # # #         self.title_label.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
# # # # #         self.title_label.setStyleSheet(f"color: {TEXT_PRIMARY}; background: transparent;")

# # # # #         self.badge_label = QLabel(badge_text)
# # # # #         self.badge_label.setAlignment(Qt.AlignCenter)
# # # # #         self.badge_label.setWordWrap(True)
# # # # #         self.badge_label.setStyleSheet(
# # # # #             f"color: {icon_color}; background-color: rgba(255,255,255,0.06); "
# # # # #             f"border: none; border-radius: 999px; padding: 4px 12px; font-size: 10.5pt;"
# # # # #         )

# # # # #         self.hint_label = QLabel(hint)
# # # # #         self.hint_label.setAlignment(Qt.AlignCenter)
# # # # #         self.hint_label.setWordWrap(True)
# # # # #         self.hint_label.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 10pt; background: transparent;")

# # # # #         self.filename_label = QLabel("")
# # # # #         self.filename_label.setAlignment(Qt.AlignCenter)
# # # # #         self.filename_label.setWordWrap(True)
# # # # #         self.filename_label.setStyleSheet(f"color: {SUCCESS}; font-size: 10pt; background: transparent;")
# # # # #         self.filename_label.hide()

# # # # #         layout.addLayout(icon_row)
# # # # #         layout.addWidget(self.title_label)
# # # # #         layout.addWidget(self.badge_label)
# # # # #         layout.addWidget(self.hint_label)
# # # # #         layout.addWidget(self.filename_label)

# # # # #         self._default_badge = badge_text
# # # # #         self._icon_color = icon_color

# # # # #     def _set_style(self, active: bool, filled: bool = False):
# # # # #         if active:
# # # # #             border, bg = BORDER_ACTIVE, ACCENT_SOFT
# # # # #         elif filled:
# # # # #             border, bg = TERTIARY, PANEL_LIGHT
# # # # #         else:
# # # # #             border, bg = BORDER, PANEL
# # # # #         self.setStyleSheet(
# # # # #             f"""
# # # # #             QFrame {{
# # # # #                 background-color: {bg};
# # # # #                 border: 1.5px solid {border};
# # # # #                 border-radius: 16px;
# # # # #             }}
# # # # #             """
# # # # #         )

# # # # #     def _valid_path(self, path: str) -> bool:
# # # # #         return path.lower().endswith(tuple(self.extensions))

# # # # #     def _apply_path(self, path: str):
# # # # #         self.path = path
# # # # #         self.badge_label.setText(os.path.basename(path))
# # # # #         self.filename_label.hide()
# # # # #         self.hint_label.setText("Click or drop again to replace")
# # # # #         self._set_style(active=False, filled=True)
# # # # #         self.on_pick()

# # # # #     def mousePressEvent(self, event):
# # # # #         filt = " ".join(f"*{e}" for e in self.extensions)
# # # # #         path, _ = QFileDialog.getOpenFileName(self, "Select file", "", f"Supported Files ({filt})")
# # # # #         if path and self._valid_path(path):
# # # # #             self._apply_path(path)

# # # # #     def dragEnterEvent(self, event: QDragEnterEvent):
# # # # #         if event.mimeData().hasUrls():
# # # # #             url = event.mimeData().urls()[0]
# # # # #             if self._valid_path(url.toLocalFile()):
# # # # #                 self._set_style(active=True)
# # # # #                 event.acceptProposedAction()
# # # # #                 return
# # # # #         event.ignore()

# # # # #     def dragLeaveEvent(self, event):
# # # # #         self._set_style(active=False, filled=bool(self.path))

# # # # #     def dropEvent(self, event: QDropEvent):
# # # # #         url = event.mimeData().urls()[0]
# # # # #         path = url.toLocalFile()
# # # # #         if self._valid_path(path):
# # # # #             self._apply_path(path)
# # # # #             event.acceptProposedAction()
# # # # #         else:
# # # # #             self._set_style(active=False, filled=bool(self.path))
# # # # #             event.ignore()


# # # # # class WorkerThread(QThread):
# # # # #     finished_ok = Signal(dict)
# # # # #     finished_err = Signal(str)

# # # # #     def __init__(self, excel_path, data_path, skip_unrelated_sheets):
# # # # #         super().__init__()
# # # # #         self.excel_path = excel_path
# # # # #         self.data_path = data_path
# # # # #         self.skip_unrelated_sheets = skip_unrelated_sheets

# # # # #     def run(self):
# # # # #         try:
# # # # #             summary = fill_workbook(
# # # # #                 self.excel_path,
# # # # #                 self.data_path,
# # # # #                 skip_unrelated_sheets=self.skip_unrelated_sheets,
# # # # #             )
# # # # #             self.finished_ok.emit(summary)
# # # # #         except PermissionError as e:
# # # # #             self.finished_err.emit(str(e))
# # # # #         except Exception:
# # # # #             self.finished_err.emit(traceback.format_exc())


# # # # # class MainWindow(QMainWindow):
# # # # #     def __init__(self):
# # # # #         super().__init__()
# # # # #         self.setWindowTitle("Excel Auto-Fill Bot")
# # # # #         self.resize(820, 760)
# # # # #         self.setMinimumSize(700, 620)
# # # # #         self.setStyleSheet(f"QMainWindow {{ background-color: {BG}; }}")
# # # # #         if os.path.exists(LOGO_PATH):
# # # # #             self.setWindowIcon(QIcon(LOGO_PATH))

# # # # #         central = QWidget()
# # # # #         self.setCentralWidget(central)
# # # # #         root = QVBoxLayout(central)
# # # # #         root.setContentsMargins(0, 0, 0, 0)
# # # # #         root.setSpacing(0)

# # # # #         # ---------- Header (AppBar) ----------
# # # # #         header = QWidget()
# # # # #         header.setStyleSheet(f"background-color: {SURFACE}; border-bottom: 1px solid {BORDER};")
# # # # #         header_layout = QVBoxLayout(header)
# # # # #         header_layout.setContentsMargins(32, 26, 32, 22)
# # # # #         header_layout.setSpacing(4)

# # # # #         header_row = QHBoxLayout()
# # # # #         badge = QLabel()
# # # # #         badge.setFixedSize(48, 48)
# # # # #         if os.path.exists(LOGO_PATH):
# # # # #             badge.setPixmap(
# # # # #                 QPixmap(LOGO_PATH).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
# # # # #             )
# # # # #         else:
# # # # #             badge.setText("⚡")
# # # # #             badge.setStyleSheet(
# # # # #                 f"background-color: {ACCENT_SOFT}; border-radius: 12px; font-size: 20px; color: {ACCENT};"
# # # # #             )
# # # # #         badge.setAlignment(Qt.AlignCenter)

# # # # #         title_box = QVBoxLayout()
# # # # #         title_box.setSpacing(2)
# # # # #         heading = QLabel("Excel Auto-Fill Bot")
# # # # #         heading.setFont(QFont("Segoe UI", 21, QFont.Bold))
# # # # #         heading.setStyleSheet(f"color: {TEXT_PRIMARY};")
# # # # #         title_box.addWidget(heading)

# # # # #         header_row.addWidget(badge)
# # # # #         header_row.addSpacing(10)
# # # # #         header_row.addLayout(title_box)
# # # # #         header_row.addStretch()
# # # # #         header_layout.addLayout(header_row)

# # # # #         root.addWidget(header)

# # # # #         # ---------- Body ----------
# # # # #         body = QWidget()
# # # # #         body_layout = QVBoxLayout(body)
# # # # #         body_layout.setContentsMargins(32, 28, 32, 28)
# # # # #         body_layout.setSpacing(20)

# # # # #         # Drop zones (bento grid)
# # # # #         picker_row = QHBoxLayout()
# # # # #         picker_row.setSpacing(16)
# # # # #         self.excel_zone = DropZone(
# # # # #             "📊", TERTIARY, "Excel File",
# # # # #             "Drag & drop .xlsx here\nor click to browse",
# # # # #             [".xlsx", ".xlsm"], "multisheet_workbook.xlsx", self._update_run_state,
# # # # #         )
# # # # #         self.data_zone = DropZone(
# # # # #             "🗂", SECONDARY, "Data File (JSON / XML)",
# # # # #             "Drag & drop .json / .xml here\nor click to browse",
# # # # #             [".json", ".xml"], "response.json", self._update_run_state,
# # # # #         )
# # # # #         picker_row.addWidget(self.excel_zone)
# # # # #         picker_row.addWidget(self.data_zone)
# # # # #         body_layout.addLayout(picker_row)

# # # # #         # Options
# # # # #         self.fill_all_checkbox = QCheckBox(
# # # # #             "  Also fill sheets that don't match this response at all (fills every cell with N/A)"
# # # # #         )
# # # # #         self.fill_all_checkbox.setChecked(False)
# # # # #         self.fill_all_checkbox.setStyleSheet(
# # # # #             f"color: {TEXT_SECONDARY}; font-size: 10pt;"
# # # # #         )
# # # # #         body_layout.addWidget(self.fill_all_checkbox)

# # # # #         # Run button
# # # # #         self.run_btn = QPushButton("▶  Run — Fill Excel File")
# # # # #         self.run_btn.setEnabled(False)
# # # # #         self.run_btn.setCursor(Qt.PointingHandCursor)
# # # # #         self.run_btn.setMinimumHeight(56)
# # # # #         self.run_btn.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
# # # # #         self.run_btn.setStyleSheet(
# # # # #             f"""
# # # # #             QPushButton {{
# # # # #                 background-color: {ACCENT_CONTAINER};
# # # # #                 color: #07006c;
# # # # #                 border-radius: 12px;
# # # # #                 border: none;
# # # # #             }}
# # # # #             QPushButton:disabled {{
# # # # #                 background-color: {PANEL_LIGHT};
# # # # #                 color: {TEXT_MUTED};
# # # # #             }}
# # # # #             QPushButton:hover:!disabled {{
# # # # #                 background-color: {ACCENT_HOVER};
# # # # #                 color: #ffffff;
# # # # #             }}
# # # # #             QPushButton:pressed:!disabled {{
# # # # #                 background-color: #4f5bd1;
# # # # #                 color: #ffffff;
# # # # #             }}
# # # # #             """
# # # # #         )
# # # # #         self.run_btn.clicked.connect(self._run)
# # # # #         body_layout.addWidget(self.run_btn)

# # # # #         # Status pill
# # # # #         self.status_label = QLabel("")
# # # # #         self.status_label.setFont(QFont("Segoe UI", 10, QFont.Medium))
# # # # #         self.status_label.setAlignment(Qt.AlignCenter)
# # # # #         self.status_label.setMinimumHeight(40)
# # # # #         self.status_label.hide()
# # # # #         body_layout.addWidget(self.status_label)

# # # # #         # Activity log (terminal style)
# # # # #         log_card = QFrame()
# # # # #         log_card.setStyleSheet(
# # # # #             f"QFrame {{ background-color: {PANEL}; border: 1px solid {BORDER}; border-radius: 14px; }}"
# # # # #         )
# # # # #         log_card_layout = QVBoxLayout(log_card)
# # # # #         log_card_layout.setContentsMargins(0, 0, 0, 0)
# # # # #         log_card_layout.setSpacing(0)

# # # # #         term_bar = QWidget()
# # # # #         term_bar.setStyleSheet(
# # # # #             f"background-color: {SURFACE_HIGH}; border-top-left-radius: 14px; border-top-right-radius: 14px;"
# # # # #         )
# # # # #         term_bar_layout = QHBoxLayout(term_bar)
# # # # #         term_bar_layout.setContentsMargins(14, 8, 14, 8)
# # # # #         for dot_color in (ERROR, SECONDARY, TERTIARY):
# # # # #             dot = QLabel()
# # # # #             dot.setFixedSize(10, 10)
# # # # #             dot.setStyleSheet(f"background-color: {dot_color}; border-radius: 5px;")
# # # # #             term_bar_layout.addWidget(dot)
# # # # #             term_bar_layout.addSpacing(4)
# # # # #         term_bar_layout.addSpacing(10)
# # # # #         term_label = QLabel("TERMINAL — PROCESS_MAPPING")
# # # # #         term_label.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 8.5pt; letter-spacing: 1px;")
# # # # #         term_bar_layout.addWidget(term_label)
# # # # #         term_bar_layout.addStretch()
# # # # #         log_card_layout.addWidget(term_bar)

# # # # #         self.log = QTextEdit()
# # # # #         self.log.setReadOnly(True)
# # # # #         self.log.setMinimumHeight(180)
# # # # #         self.log.setFont(QFont("Consolas", 10))
# # # # #         self.log.setStyleSheet(
# # # # #             f"""
# # # # #             QTextEdit {{
# # # # #                 background-color: transparent;
# # # # #                 color: {TEXT_SECONDARY};
# # # # #                 border: none;
# # # # #                 padding: 16px;
# # # # #             }}
# # # # #             """
# # # # #         )
# # # # #         log_card_layout.addWidget(self.log, stretch=1)

# # # # #         body_layout.addWidget(log_card, stretch=1)
# # # # #         root.addWidget(body, stretch=1)

# # # # #         self.worker = None
# # # # #         self._log_placeholder()

# # # # #     # ------------------------------------------------------------------
# # # # #     def _log_placeholder(self):
# # # # #         self.log.setHtml(
# # # # #             f"<span style='color:{TEXT_MUTED};'>Waiting for files… select or drop an Excel file "
# # # # #             f"and a JSON/XML data file above, then hit Run.</span>"
# # # # #         )

# # # # #     def _update_run_state(self):
# # # # #         self.run_btn.setEnabled(bool(self.excel_zone.path and self.data_zone.path))

# # # # #     def _set_status(self, text, kind="info"):
# # # # #         color = {"success": SUCCESS, "error": ERROR, "info": ACCENT}[kind]
# # # # #         self.status_label.setText(f"  {text}")
# # # # #         self.status_label.setStyleSheet(
# # # # #             f"color: {color}; background-color: rgba(78,222,163,0.08); border: 1px solid {color}; "
# # # # #             f"border-radius: 10px; padding: 10px;"
# # # # #         )
# # # # #         self.status_label.show()

# # # # #     def _run(self):
# # # # #         self.run_btn.setEnabled(False)
# # # # #         self.run_btn.setText("Processing…")
# # # # #         self.status_label.hide()
# # # # #         self.log.clear()
# # # # #         self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Excel:</b> {self.excel_zone.path}")
# # # # #         self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Data:</b> {self.data_zone.path}")
# # # # #         self._append_log(f"<span style='color:{TEXT_MUTED}'>Reading sheets, detecting headers, matching data…</span><br>")

# # # # #         skip_unrelated = not self.fill_all_checkbox.isChecked()
# # # # #         self.worker = WorkerThread(self.excel_zone.path, self.data_zone.path, skip_unrelated)
# # # # #         self.worker.finished_ok.connect(self._on_success)
# # # # #         self.worker.finished_err.connect(self._on_error)
# # # # #         self.worker.start()

# # # # #     def _append_log(self, html):
# # # # #         self.log.append(html)

# # # # #     def _on_success(self, summary):
# # # # #         self.run_btn.setEnabled(True)
# # # # #         self.run_btn.setText("▶  Run — Fill Excel File")

# # # # #         self._append_log(
# # # # #             f"<span style='color:{TEXT_PRIMARY}'>Records found in data file: "
# # # # #             f"<b>{summary['records_processed']}</b></span><br>"
# # # # #         )
# # # # #         for sheet in summary["sheets"]:
# # # # #             self._append_log(f"<b style='color:{SECONDARY}'>■ {sheet['sheet']}</b>")
# # # # #             if sheet["skipped"]:
# # # # #                 self._append_log(
# # # # #                     f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;No matching headers — sheet left untouched.</span><br>"
# # # # #                 )
# # # # #                 continue
# # # # #             self._append_log(
# # # # #                 f"<span style='color:{TERTIARY}'>&nbsp;&nbsp;Rows filled: {sheet['rows_filled']} "
# # # # #                 f"(new: {sheet['rows_added']}, replaced duplicates: {sheet['rows_replaced']})</span>"
# # # # #             )
# # # # #             if sheet["matched_headers"]:
# # # # #                 self._append_log(f"<span style='color:{TEXT_SECONDARY}'>&nbsp;&nbsp;Matched columns:</span>")
# # # # #                 for header, key in sheet["matched_headers"].items():
# # # # #                     self._append_log(
# # # # #                         f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;&nbsp;&nbsp;'{header}' ← '{key}'</span>"
# # # # #                     )
# # # # #             if sheet["unmatched_headers"]:
# # # # #                 self._append_log(
# # # # #                     f"<span style='color:{ERROR}'>&nbsp;&nbsp;Filled with N/A (no matching data): "
# # # # #                     f"{', '.join(sheet['unmatched_headers'])}</span>"
# # # # #                 )
# # # # #             self._append_log("")

# # # # #         if summary.get("save_fallback"):
# # # # #             self._append_log(
# # # # #                 f"<b style='color:{ERROR}'>⚠ '{self.excel_zone.path}' was locked (probably open in Excel) — "
# # # # #                 f"saved a copy instead:</b>"
# # # # #             )
# # # # #             self._append_log(f"<b style='color:{TERTIARY}'>✓ Saved to: {summary['saved_to']}</b>")
# # # # #             self._set_status(
# # # # #                 "⚠ Original file was open/locked — saved a copy instead (see log for path).",
# # # # #                 kind="error",
# # # # #             )
# # # # #         else:
# # # # #             self._append_log(f"<b style='color:{TERTIARY}'>✓ Done — Excel file updated in place.</b>")
# # # # #             self._set_status("✓ Excel file filled and saved successfully.", kind="success")

# # # # #     def _on_error(self, err_text):
# # # # #         self.run_btn.setEnabled(True)
# # # # #         self.run_btn.setText("▶  Run — Fill Excel File")
# # # # #         self._append_log(f"<span style='color:{ERROR}'><b>ERROR</b></span>")
# # # # #         self._append_log(f"<span style='color:{ERROR}'>{err_text}</span>")
# # # # #         self._set_status("✗ Something went wrong — see log for details.", kind="error")


# # # # # def main():
# # # # #     app = QApplication(sys.argv)
# # # # #     app.setStyle("Fusion")
# # # # #     if os.path.exists(LOGO_PATH):
# # # # #         app.setWindowIcon(QIcon(LOGO_PATH))
# # # # #     win = MainWindow()
# # # # #     win.show()
# # # # #     sys.exit(app.exec())


# # # # # if __name__ == "__main__":
# # # # #     main()











# # """
# # Excel Auto-Fill Bot - Desktop App

# # Pick (or drag & drop) an Excel file (multi-sheet, dynamic headers - nothing
# # hardcoded) and a JSON or XML data file (e.g. an Aadhaar-style API response).
# # The app fuzzy-matches the data's keys against each sheet's headers, fills
# # matching values into a brand new row of every sheet, writes "N/A" for
# # anything with no matching data, and overwrites the same Excel file.

# # - Har record (chahe duplicate hi kyu na ho) apni alag row leta hai - kuch bhi
# #   skip/merge nahi hota.
# # - By default koi bhi sheet skip nahi hoti - jis sheet mein match nahi milta
# #   wahan bhi ek row jaake N/A se bhar jaati hai. "Sirf unrelated sheets skip
# #   karo" checkbox se old (strict) behaviour wapas mil jata hai.
# # """

# # import os
# # import sys
# # import traceback

# # from PySide6.QtCore import Qt, QThread, Signal
# # from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent, QColor, QIcon, QPixmap
# # from PySide6.QtWidgets import (
# #     QApplication,
# #     QCheckBox,
# #     QFileDialog,
# #     QFrame,
# #     QGraphicsDropShadowEffect,
# #     QHBoxLayout,
# #     QLabel,
# #     QMainWindow,
# #     QPushButton,
# #     QSizePolicy,
# #     QTextEdit,
# #     QVBoxLayout,
# #     QWidget,
# # )

# # from filler_core import fill_workbook

# # # ---------------------------------------------------------------------------
# # # Palette (matches the Excel Auto-Fill Bot design spec)
# # # ---------------------------------------------------------------------------
# # BG = "#0b1326"
# # SURFACE = "#0b1326"
# # PANEL = "#171f33"
# # PANEL_LIGHT = "#1e2338"
# # SURFACE_HIGH = "#222a3d"
# # SURFACE_HIGHEST = "#2d3449"
# # BORDER = "#464554"
# # BORDER_ACTIVE = "#8083ff"
# # ACCENT = "#c0c1ff"                 # primary
# # ACCENT_CONTAINER = "#8083ff"       # primary-container
# # ACCENT_HOVER = "#6f72e8"
# # ACCENT_SOFT = "#2a2d55"
# # SECONDARY = "#89ceff"
# # SECONDARY_CONTAINER = "#00a2e6"
# # TERTIARY = "#4edea3"
# # TERTIARY_CONTAINER = "#00885d"
# # TEXT_PRIMARY = "#dae2fd"
# # TEXT_SECONDARY = "#c7c4d7"
# # TEXT_MUTED = "#908fa0"
# # SUCCESS = "#4edea3"
# # ERROR = "#ffb4ab"

# # ASSET_DIR = os.path.dirname(os.path.abspath(__file__))
# # LOGO_PATH = os.path.join(ASSET_DIR, "logo.png")


# # class DropZone(QFrame):
# #     """Clickable + drag-and-drop file picker box (bento-grid card)."""

# #     def __init__(self, icon, icon_color, title, hint, extensions, badge_text, on_pick):
# #         super().__init__()
# #         self.extensions = extensions
# #         self.on_pick = on_pick
# #         self.path = None

# #         self.setAcceptDrops(True)
# #         self.setCursor(Qt.PointingHandCursor)
# #         self.setMinimumHeight(170)
# #         self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
# #         self._set_style(active=False)

# #         shadow = QGraphicsDropShadowEffect(self)
# #         shadow.setBlurRadius(28)
# #         shadow.setOffset(0, 8)
# #         shadow.setColor(QColor(0, 0, 0, 110))
# #         self.setGraphicsEffect(shadow)

# #         layout = QVBoxLayout(self)
# #         layout.setContentsMargins(20, 24, 20, 24)
# #         layout.setSpacing(8)
# #         layout.setAlignment(Qt.AlignCenter)

# #         icon_wrap = QLabel(icon)
# #         icon_wrap.setAlignment(Qt.AlignCenter)
# #         icon_wrap.setFixedSize(64, 64)
# #         icon_wrap.setStyleSheet(
# #             f"font-size: 26px; background-color: {SURFACE_HIGH}; border: none; border-radius: 18px; color: {icon_color};"
# #         )

# #         icon_row = QHBoxLayout()
# #         icon_row.addStretch()
# #         icon_row.addWidget(icon_wrap)
# #         icon_row.addStretch()

# #         self.title_label = QLabel(title)
# #         self.title_label.setAlignment(Qt.AlignCenter)
# #         self.title_label.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
# #         self.title_label.setStyleSheet(f"""
# # QLabel {{
# #     color: {TEXT_PRIMARY};
# #     background: transparent;
# #     border: none;
# # }}
# # """)

# #         self.badge_label = QLabel(badge_text)
# #         self.badge_label.setAlignment(Qt.AlignCenter)
# #         self.badge_label.setWordWrap(True)
# #         self.badge_label.setStyleSheet(
# #             f"color: {icon_color}; background-color: rgba(255,255,255,0.06); "
# #             f"border: none; border-radius: 999px; padding: 4px 12px; font-size: 10.5pt;"
# #         )

# #         self.hint_label = QLabel(hint)
# #         self.hint_label.setAlignment(Qt.AlignCenter)
# #         self.hint_label.setWordWrap(True)
# #         self.hint_label.setStyleSheet(f"""
# # QLabel {{
# #  color: {TEXT_MUTED};
# #  background: transparent;
# #  border: none;
# #  font-size:10pt;
# # }}
# # """)

# #         self.filename_label = QLabel("")
# #         self.filename_label.setAlignment(Qt.AlignCenter)
# #         self.filename_label.setWordWrap(True)
# #         self.filename_label.setStyleSheet(f"""
# # QLabel {{
# #  color:{SUCCESS};
# #  background:transparent;
# #  border:none;
# #  font-size:10pt;
# # }}
# # """)
# #         self.filename_label.hide()

# #         layout.addLayout(icon_row)
# #         layout.addWidget(self.title_label)
# #         layout.addWidget(self.badge_label)
# #         layout.addWidget(self.hint_label)
# #         layout.addWidget(self.filename_label)

# #         self._default_badge = badge_text
# #         self._icon_color = icon_color

# #     def _set_style(self, active: bool, filled: bool = False):
# #         if active:
# #             border, bg = BORDER_ACTIVE, ACCENT_SOFT
# #         elif filled:
# #             border, bg = TERTIARY, PANEL_LIGHT
# #         else:
# #             border, bg = BORDER, PANEL
# #         self.setStyleSheet(f"""
# # DropZone {{
# #     background-color: {bg};
# #     border: 2px solid {border};
# #     border-radius: 18px;
# # }}
# # QLabel {{
# #     border:none;
# #     background:transparent;
# # }}
# # """)

# #     def _valid_path(self, path: str) -> bool:
# #         return path.lower().endswith(tuple(self.extensions))

# #     def _apply_path(self, path: str):
# #         self.path = path
# #         self.badge_label.setText(os.path.basename(path))
# #         self.filename_label.hide()
# #         self.hint_label.setText("Click or drop again to replace")
# #         self._set_style(active=False, filled=True)
# #         self.on_pick()

# #     def mousePressEvent(self, event):
# #         filt = " ".join(f"*{e}" for e in self.extensions)
# #         path, _ = QFileDialog.getOpenFileName(self, "Select file", "", f"Supported Files ({filt})")
# #         if path and self._valid_path(path):
# #             self._apply_path(path)

# #     def dragEnterEvent(self, event: QDragEnterEvent):
# #         if event.mimeData().hasUrls():
# #             url = event.mimeData().urls()[0]
# #             if self._valid_path(url.toLocalFile()):
# #                 self._set_style(active=True)
# #                 event.acceptProposedAction()
# #                 return
# #         event.ignore()

# #     def dragLeaveEvent(self, event):
# #         self._set_style(active=False, filled=bool(self.path))

# #     def dropEvent(self, event: QDropEvent):
# #         url = event.mimeData().urls()[0]
# #         path = url.toLocalFile()
# #         if self._valid_path(path):
# #             self._apply_path(path)
# #             event.acceptProposedAction()
# #         else:
# #             self._set_style(active=False, filled=bool(self.path))
# #             event.ignore()


# # class WorkerThread(QThread):
# #     finished_ok = Signal(dict)
# #     finished_err = Signal(str)

# #     def __init__(self, excel_path, data_path, skip_unrelated_sheets):
# #         super().__init__()
# #         self.excel_path = excel_path
# #         self.data_path = data_path
# #         self.skip_unrelated_sheets = skip_unrelated_sheets

# #     def run(self):
# #         try:
# #             summary = fill_workbook(
# #                 self.excel_path,
# #                 self.data_path,
# #                 skip_unrelated_sheets=self.skip_unrelated_sheets,
# #             )
# #             self.finished_ok.emit(summary)
# #         except PermissionError as e:
# #             self.finished_err.emit(str(e))
# #         except Exception:
# #             self.finished_err.emit(traceback.format_exc())


# # class MainWindow(QMainWindow):
# #     def __init__(self):
# #         super().__init__()
# #         self.setWindowTitle("Excel Auto-Fill Bot")
# #         self.resize(820, 760)
# #         self.setMinimumSize(700, 620)
# #         self.setStyleSheet(f"QMainWindow {{ background-color: {BG}; }}")
# #         if os.path.exists(LOGO_PATH):
# #             self.setWindowIcon(QIcon(LOGO_PATH))

# #         central = QWidget()
# #         self.setCentralWidget(central)
# #         root = QVBoxLayout(central)
# #         root.setContentsMargins(0, 0, 0, 0)
# #         root.setSpacing(0)

# #         # ---------- Header (AppBar) ----------
# #         header = QWidget()
# #         header.setStyleSheet(f"background-color: {SURFACE}; border-bottom: 1px solid {BORDER};")
# #         header_layout = QVBoxLayout(header)
# #         header_layout.setContentsMargins(32, 26, 32, 22)
# #         header_layout.setSpacing(4)

# #         header_row = QHBoxLayout()
# #         badge = QLabel()
# #         badge.setFixedSize(48, 48)
# #         if os.path.exists(LOGO_PATH):
# #             badge.setPixmap(
# #                 QPixmap(LOGO_PATH).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
# #             )
# #         else:
# #             badge.setText("⚡")
# #             badge.setStyleSheet(
# #                 f"background-color: {ACCENT_SOFT}; border-radius: 12px; font-size: 20px; color: {ACCENT};"
# #             )
# #         badge.setAlignment(Qt.AlignCenter)

# #         title_box = QVBoxLayout()
# #         title_box.setSpacing(2)
# #         heading = QLabel("Excel Auto-Fill Bot")
# #         heading.setFont(QFont("Segoe UI", 21, QFont.Bold))
# #         heading.setStyleSheet(f"color: {TEXT_PRIMARY};")
# #         title_box.addWidget(heading)

# #         header_row.addWidget(badge)
# #         header_row.addSpacing(10)
# #         header_row.addLayout(title_box)
# #         header_row.addStretch()
# #         header_layout.addLayout(header_row)

# #         root.addWidget(header)

# #         # ---------- Body ----------
# #         body = QWidget()
# #         body_layout = QVBoxLayout(body)
# #         body_layout.setContentsMargins(32, 28, 32, 28)
# #         body_layout.setSpacing(20)

# #         # Drop zones (bento grid)
# #         picker_row = QHBoxLayout()
# #         picker_row.setSpacing(16)
# #         self.excel_zone = DropZone(
# #             "📊", TERTIARY, "Excel File",
# #             "Drag & drop .xlsx here\nor click to browse",
# #             [".xlsx", ".xlsm"], "multisheet_workbook.xlsx", self._update_run_state,
# #         )
# #         self.data_zone = DropZone(
# #             "🗂", SECONDARY, "Data File (JSON / XML)",
# #             "Drag & drop .json / .xml here\nor click to browse",
# #             [".json", ".xml"], "response.json", self._update_run_state,
# #         )
# #         picker_row.addWidget(self.excel_zone)
# #         picker_row.addWidget(self.data_zone)
# #         body_layout.addLayout(picker_row)

# #         # Options
# #         self.fill_all_checkbox = QCheckBox(
# #             "  Also fill sheets that don't match this response at all (fills every cell with N/A)"
# #         )
# #         self.fill_all_checkbox.setChecked(False)
# #         self.fill_all_checkbox.setStyleSheet(
# #             f"color: {TEXT_SECONDARY}; font-size: 10pt;"
# #         )
# #         body_layout.addWidget(self.fill_all_checkbox)

# #         # Run button
# #         self.run_btn = QPushButton("▶  Run — Fill Excel File")
# #         self.run_btn.setEnabled(False)
# #         self.run_btn.setCursor(Qt.PointingHandCursor)
# #         self.run_btn.setMinimumHeight(56)
# #         self.run_btn.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
# #         self.run_btn.setStyleSheet(
# #             f"""
# #             QPushButton {{
# #                 background-color: {ACCENT_CONTAINER};
# #                 color: #07006c;
# #                 border-radius: 12px;
# #                 border: none;
# #             }}
# #             QPushButton:disabled {{
# #                 background-color: {PANEL_LIGHT};
# #                 color: {TEXT_MUTED};
# #             }}
# #             QPushButton:hover:!disabled {{
# #                 background-color: {ACCENT_HOVER};
# #                 color: #ffffff;
# #             }}
# #             QPushButton:pressed:!disabled {{
# #                 background-color: #4f5bd1;
# #                 color: #ffffff;
# #             }}
# #             """
# #         )
# #         self.run_btn.clicked.connect(self._run)
# #         body_layout.addWidget(self.run_btn)

# #         # Status pill
# #         self.status_label = QLabel("")
# #         self.status_label.setFont(QFont("Segoe UI", 10, QFont.Medium))
# #         self.status_label.setAlignment(Qt.AlignCenter)
# #         self.status_label.setMinimumHeight(40)
# #         self.status_label.hide()
# #         body_layout.addWidget(self.status_label)

# #         # Activity log (terminal style)
# #         log_card = QFrame()
# #         log_card.setStyleSheet(
# #             f"QFrame {{ background-color: {PANEL}; border: 1px solid {BORDER}; border-radius: 14px; }}"
# #         )
# #         log_card_layout = QVBoxLayout(log_card)
# #         log_card_layout.setContentsMargins(0, 0, 0, 0)
# #         log_card_layout.setSpacing(0)

# #         term_bar = QWidget()
# #         term_bar.setStyleSheet(
# #             f"background-color: {SURFACE_HIGH}; border-top-left-radius: 14px; border-top-right-radius: 14px;"
# #         )
# #         term_bar_layout = QHBoxLayout(term_bar)
# #         term_bar_layout.setContentsMargins(14, 8, 14, 8)
# #         for dot_color in (ERROR, SECONDARY, TERTIARY):
# #             dot = QLabel()
# #             dot.setFixedSize(10, 10)
# #             dot.setStyleSheet(f"background-color: {dot_color}; border-radius: 5px;")
# #             term_bar_layout.addWidget(dot)
# #             term_bar_layout.addSpacing(4)
# #         term_bar_layout.addSpacing(10)
# #         term_label = QLabel("TERMINAL — PROCESS_MAPPING")
# #         term_label.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 8.5pt; letter-spacing: 1px;")
# #         term_bar_layout.addWidget(term_label)
# #         term_bar_layout.addStretch()
# #         log_card_layout.addWidget(term_bar)

# #         self.log = QTextEdit()
# #         self.log.setReadOnly(True)
# #         self.log.setMinimumHeight(180)
# #         self.log.setFont(QFont("Consolas", 10))
# #         self.log.setStyleSheet(
# #             f"""
# #             QTextEdit {{
# #                 background-color: transparent;
# #                 color: {TEXT_SECONDARY};
# #                 border: none;
# #                 padding: 16px;
# #             }}
# #             """
# #         )
# #         log_card_layout.addWidget(self.log, stretch=1)

# #         body_layout.addWidget(log_card, stretch=1)
# #         root.addWidget(body, stretch=1)

# #         self.worker = None
# #         self._log_placeholder()

# #     # ------------------------------------------------------------------
# #     def _log_placeholder(self):
# #         self.log.setHtml(
# #             f"<span style='color:{TEXT_MUTED};'>Waiting for files… select or drop an Excel file "
# #             f"and a JSON/XML data file above, then hit Run.</span>"
# #         )

# #     def _update_run_state(self):
# #         self.run_btn.setEnabled(bool(self.excel_zone.path and self.data_zone.path))

# #     def _set_status(self, text, kind="info"):
# #         color = {"success": SUCCESS, "error": ERROR, "info": ACCENT}[kind]
# #         self.status_label.setText(f"  {text}")
# #         self.status_label.setStyleSheet(
# #             f"color: {color}; background-color: rgba(78,222,163,0.08); border: 1px solid {color}; "
# #             f"border-radius: 10px; padding: 10px;"
# #         )
# #         self.status_label.show()

# #     def _run(self):
# #         self.run_btn.setEnabled(False)
# #         self.run_btn.setText("Processing…")
# #         self.status_label.hide()
# #         self.log.clear()
# #         self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Excel:</b> {self.excel_zone.path}")
# #         self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Data:</b> {self.data_zone.path}")
# #         self._append_log(f"<span style='color:{TEXT_MUTED}'>Reading sheets, detecting headers, matching data…</span><br>")

# #         skip_unrelated = not self.fill_all_checkbox.isChecked()
# #         self.worker = WorkerThread(self.excel_zone.path, self.data_zone.path, skip_unrelated)
# #         self.worker.finished_ok.connect(self._on_success)
# #         self.worker.finished_err.connect(self._on_error)
# #         self.worker.start()

# #     def _append_log(self, html):
# #         self.log.append(html)

# #     def _on_success(self, summary):
# #         self.run_btn.setEnabled(True)
# #         self.run_btn.setText("▶  Run — Fill Excel File")

# #         self._append_log(
# #             f"<span style='color:{TEXT_PRIMARY}'>Records found in data file: "
# #             f"<b>{summary['records_processed']}</b></span><br>"
# #         )
# #         for sheet in summary["sheets"]:
# #             self._append_log(f"<b style='color:{SECONDARY}'>■ {sheet['sheet']}</b>")
# #             if sheet["skipped"]:
# #                 self._append_log(
# #                     f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;No matching headers — sheet left untouched.</span><br>"
# #                 )
# #                 continue
# #             self._append_log(
# #                 f"<span style='color:{TERTIARY}'>&nbsp;&nbsp;Rows filled: {sheet['rows_filled']} "
# #                 f"(new: {sheet['rows_added']}, replaced duplicates: {sheet['rows_replaced']})</span>"
# #             )
# #             if sheet["matched_headers"]:
# #                 self._append_log(f"<span style='color:{TEXT_SECONDARY}'>&nbsp;&nbsp;Matched columns:</span>")
# #                 for header, key in sheet["matched_headers"].items():
# #                     self._append_log(
# #                         f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;&nbsp;&nbsp;'{header}' ← '{key}'</span>"
# #                     )
# #             if sheet["unmatched_headers"]:
# #                 self._append_log(
# #                     f"<span style='color:{ERROR}'>&nbsp;&nbsp;Filled with N/A (no matching data): "
# #                     f"{', '.join(sheet['unmatched_headers'])}</span>"
# #                 )
# #             self._append_log("")

# #         if summary.get("save_fallback"):
# #             self._append_log(
# #                 f"<b style='color:{ERROR}'>⚠ '{self.excel_zone.path}' was locked (probably open in Excel) — "
# #                 f"saved a copy instead:</b>"
# #             )
# #             self._append_log(f"<b style='color:{TERTIARY}'>✓ Saved to: {summary['saved_to']}</b>")
# #             self._set_status(
# #                 "⚠ Original file was open/locked — saved a copy instead (see log for path).",
# #                 kind="error",
# #             )
# #         else:
# #             self._append_log(f"<b style='color:{TERTIARY}'>✓ Done — Excel file updated in place.</b>")
# #             self._set_status("✓ Excel file filled and saved successfully.", kind="success")

# #     def _on_error(self, err_text):
# #         self.run_btn.setEnabled(True)
# #         self.run_btn.setText("▶  Run — Fill Excel File")
# #         self._append_log(f"<span style='color:{ERROR}'><b>ERROR</b></span>")
# #         self._append_log(f"<span style='color:{ERROR}'>{err_text}</span>")
# #         self._set_status("✗ Something went wrong — see log for details.", kind="error")


# # def main():
# #     app = QApplication(sys.argv)
# #     app.setStyle("Fusion")
# #     if os.path.exists(LOGO_PATH):
# #         app.setWindowIcon(QIcon(LOGO_PATH))
# #     win = MainWindow()
# #     win.show()
# #     sys.exit(app.exec())


# # if __name__ == "__main__":
# #     main()






















# # # """
# # # Excel Auto-Fill Bot - Desktop App

# # # Pick (or drag & drop) an Excel file (multi-sheet, dynamic headers - nothing
# # # hardcoded) and a JSON or XML data file (e.g. an Aadhaar-style API response).
# # # The app fuzzy-matches the data's keys against each sheet's headers, fills
# # # matching values into a brand new row of every sheet, writes "N/A" for
# # # anything with no matching data, and overwrites the same Excel file.

# # # - Har record (chahe duplicate hi kyu na ho) apni alag row leta hai - kuch bhi
# # #   skip/merge nahi hota.
# # # - By default koi bhi sheet skip nahi hoti - jis sheet mein match nahi milta
# # #   wahan bhi ek row jaake N/A se bhar jaati hai. "Sirf unrelated sheets skip
# # #   karo" checkbox se old (strict) behaviour wapas mil jata hai.
# # # """

# # # import os
# # # import sys
# # # import traceback

# # # import qtawesome as qta

# # # from PySide6.QtCore import (
# # #     QEasingCurve,
# # #     QPropertyAnimation,
# # #     Qt,
# # #     QThread,
# # #     QTimer,
# # #     Signal,
# # # )
# # # from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent, QColor, QIcon, QPixmap
# # # from PySide6.QtWidgets import (
# # #     QApplication,
# # #     QCheckBox,
# # #     QFileDialog,
# # #     QFrame,
# # #     QGraphicsDropShadowEffect,
# # #     QGraphicsOpacityEffect,
# # #     QHBoxLayout,
# # #     QLabel,
# # #     QMainWindow,
# # #     QPushButton,
# # #     QSizePolicy,
# # #     QTextEdit,
# # #     QVBoxLayout,
# # #     QWidget,
# # # )

# # # from filler_core import fill_workbook

# # # # ═══════════════════════════════════════════════════════════════════════════════
# # # # Design Tokens
# # # # ═══════════════════════════════════════════════════════════════════════════════

# # # # Deep backgrounds
# # # BG = "#09090b"
# # # SURFACE = "#0c0c10"
# # # PANEL = "#13131a"
# # # PANEL_LIGHT = "#1a1a23"
# # # SURFACE_HIGH = "#22222d"
# # # SURFACE_HIGHEST = "#2c2c3a"

# # # # Borders
# # # BORDER = "rgba(255, 255, 255, 0.06)"
# # # BORDER_ACTIVE = "#6366f1"

# # # # Accent — indigo-violet
# # # ACCENT = "#b4b8ff"
# # # ACCENT_CONTAINER = "#6366f1"
# # # ACCENT_HOVER = "#818cf8"
# # # ACCENT_SOFT = "rgba(99, 102, 241, 0.10)"

# # # # Secondary — cyan
# # # SECONDARY = "#67e8f9"
# # # SECONDARY_CONTAINER = "#06b6d4"

# # # # Tertiary — emerald
# # # TERTIARY = "#34d399"
# # # TERTIARY_CONTAINER = "#10b981"

# # # # Typography hierarchy
# # # TEXT_PRIMARY = "#f4f4f5"
# # # TEXT_SECONDARY = "#a1a1aa"
# # # TEXT_MUTED = "#71717a"

# # # # Semantic
# # # SUCCESS = "#34d399"
# # # ERROR = "#f87171"
# # # WARNING = "#fbbf24"

# # # # Assets
# # # ASSET_DIR = os.path.dirname(os.path.abspath(__file__))
# # # LOGO_PATH = os.path.join(ASSET_DIR, "logo.png")

# # # # Animation constants
# # # _ANIM_FAST = 180
# # # _ANIM_NORMAL = 280
# # # _ANIM_SLOW = 400
# # # _SHADOW_DEFAULT = 40
# # # _SHADOW_HOVER = 56
# # # _SHADOW_DRAG = 72

# # # # Layout constants
# # # _BODY_MARGIN_H = 36
# # # _BODY_MARGIN_V = 32
# # # _SECTION_SPACING = 24
# # # _CARD_RADIUS = 16
# # # _CARD_PADDING_H = 24
# # # _CARD_PADDING_V = 28
# # # _ICON_SIZE = 72
# # # _ICON_PADDING = 20
# # # _BADGE_RADIUS = 999
# # # _INDICATOR_SIZE = 18


# # # # ═══════════════════════════════════════════════════════════════════════════════
# # # # Style Helpers
# # # # ═══════════════════════════════════════════════════════════════════════════════


# # # def _make_shadow(parent, blur_radius=_SHADOW_DEFAULT, offset_y=4, alpha=60):
# # #     shadow = QGraphicsDropShadowEffect(parent)
# # #     shadow.setBlurRadius(blur_radius)
# # #     shadow.setOffset(0, offset_y)
# # #     shadow.setColor(QColor(0, 0, 0, alpha))
# # #     return shadow


# # # def _card_stylesheet(bg=PANEL, border=BORDER, radius=_CARD_RADIUS):
# # #     return (
# # #         f"QFrame {{ background-color: {bg}; border: 1px solid {border}; "
# # #         f"border-radius: {radius}px; }}"
# # #     )


# # # def _primary_button_stylesheet():
# # #     return f"""
# # #         QPushButton {{
# # #             background-color: {ACCENT_CONTAINER};
# # #             color: #ffffff;
# # #             border-radius: 12px;
# # #             border: none;
# # #             font-weight: 600;
# # #             padding: 0px 28px;
# # #         }}
# # #         QPushButton:disabled {{
# # #             background-color: rgba(255, 255, 255, 0.04);
# # #             color: {TEXT_MUTED};
# # #             border: none;
# # #         }}
# # #         QPushButton:hover:!disabled {{
# # #             background-color: {ACCENT_HOVER};
# # #             color: #ffffff;
# # #         }}
# # #         QPushButton:pressed:!disabled {{
# # #             background-color: #4f46e5;
# # #             color: #ffffff;
# # #             padding: 1px 27px 0px 29px;
# # #         }}
# # #     """


# # # def _status_pill_stylesheet(color, kind="info"):
# # #     bg_map = {
# # #         "success": "rgba(52, 211, 153, 0.08)",
# # #         "error": "rgba(248, 113, 113, 0.08)",
# # #         "info": "rgba(99, 102, 241, 0.08)",
# # #     }
# # #     bg = bg_map.get(kind, bg_map["info"])
# # #     return (
# # #         f"color: {color}; background-color: {bg}; "
# # #         f"border: 1px solid {color}; border-radius: 10px; padding: 10px 16px;"
# # #     )


# # # def _dropzone_stylesheet(active=False, filled=False):
# # #     if active:
# # #         border, bg = BORDER_ACTIVE, ACCENT_SOFT
# # #     elif filled:
# # #         border, bg = TERTIARY, PANEL_LIGHT
# # #     else:
# # #         border, bg = BORDER, PANEL
# # #     return (
# # #         f"QFrame {{ background-color: {bg}; border: 1.5px solid {border}; "
# # #         f"border-radius: {_CARD_RADIUS}px; }}"
# # #     )


# # # # ═══════════════════════════════════════════════════════════════════════════════
# # # # Reusable Widgets
# # # # ═══════════════════════════════════════════════════════════════════════════════


# # # class PrimaryButton(QPushButton):
# # #     """Styled primary-action button with hover glow animation."""

# # #     def __init__(self, text, parent=None):
# # #         super().__init__(text, parent)
# # #         self.setEnabled(False)
# # #         self.setCursor(Qt.PointingHandCursor)
# # #         self.setMinimumHeight(52)
# # #         self.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
# # #         self.setStyleSheet(_primary_button_stylesheet())
# # #         self.setShortcut("Return")
# # #         self.setToolTip("Fill the Excel workbook with matched data")

# # #         self._glow = QGraphicsDropShadowEffect(self)
# # #         self._glow.setBlurRadius(0)
# # #         self._glow.setOffset(0, 6)
# # #         glow_color = QColor(ACCENT_CONTAINER)
# # #         glow_color.setAlpha(0)
# # #         self._glow.setColor(glow_color)
# # #         self.setGraphicsEffect(self._glow)

# # #         self._glow_anim = QPropertyAnimation(self._glow, b"blurRadius")
# # #         self._glow_anim.setDuration(_ANIM_FAST)
# # #         self._glow_anim.setEasingCurve(QEasingCurve.OutCubic)

# # #         self._glow_color_anim = QPropertyAnimation(self._glow, b"color")
# # #         self._glow_color_anim.setDuration(_ANIM_FAST)
# # #         self._glow_color_anim.setEasingCurve(QEasingCurve.OutCubic)

# # #     def enterEvent(self, event):
# # #         if not self.isEnabled():
# # #             return super().enterEvent(event)
# # #         self._glow_anim.stop()
# # #         self._glow_anim.setStartValue(self._glow.blurRadius())
# # #         self._glow_anim.setEndValue(28)
# # #         self._glow_anim.start()
# # #         self._glow_color_anim.stop()
# # #         self._glow_color_anim.setStartValue(self._glow.color())
# # #         glow_color = QColor(ACCENT_CONTAINER)
# # #         glow_color.setAlpha(90)
# # #         self._glow_color_anim.setEndValue(glow_color)
# # #         self._glow_color_anim.start()
# # #         super().enterEvent(event)

# # #     def leaveEvent(self, event):
# # #         self._glow_anim.stop()
# # #         self._glow_anim.setStartValue(self._glow.blurRadius())
# # #         self._glow_anim.setEndValue(0)
# # #         self._glow_anim.start()
# # #         self._glow_color_anim.stop()
# # #         self._glow_color_anim.setStartValue(self._glow.color())
# # #         glow_color = QColor(ACCENT_CONTAINER)
# # #         glow_color.setAlpha(0)
# # #         self._glow_color_anim.setEndValue(glow_color)
# # #         self._glow_color_anim.start()
# # #         super().leaveEvent(event)


# # # class StatusPill(QLabel):
# # #     """Self-contained status pill with fade-in animation."""

# # #     _COLORS = {"success": SUCCESS, "error": ERROR, "info": ACCENT}

# # #     def __init__(self, parent=None):
# # #         super().__init__(parent)
# # #         self.setFont(QFont("Segoe UI", 10, QFont.Medium))
# # #         self.setAlignment(Qt.AlignCenter)
# # #         self.setMinimumHeight(40)

# # #         self._opacity = QGraphicsOpacityEffect(self)
# # #         self._opacity.setOpacity(0.0)
# # #         self.setGraphicsEffect(self._opacity)

# # #         self._fade = QPropertyAnimation(self._opacity, b"opacity")
# # #         self._fade.setDuration(_ANIM_SLOW)
# # #         self._fade.setEasingCurve(QEasingCurve.OutCubic)

# # #         self.hide()

# # #     def set_status(self, text, kind="info"):
# # #         color = self._COLORS[kind]
# # #         self.setText(f"  {text}")
# # #         self.setStyleSheet(_status_pill_stylesheet(color, kind))
# # #         self._opacity.setOpacity(0.0)
# # #         self.show()
# # #         self._fade.stop()
# # #         self._fade.setStartValue(0.0)
# # #         self._fade.setEndValue(1.0)
# # #         self._fade.start()

# # #     def hide(self):
# # #         self._fade.stop()
# # #         self._opacity.setOpacity(0.0)
# # #         super().hide()


# # # class TerminalWidget(QFrame):
# # #     """Terminal-style log card with processing dots animation."""

# # #     _DEFAULT_TITLE = "TERMINAL \u2014 PROCESS_MAPPING"

# # #     def __init__(self, parent=None):
# # #         super().__init__(parent)
# # #         self.setStyleSheet(_card_stylesheet(bg="#0d0d12", border=BORDER))

# # #         layout = QVBoxLayout(self)
# # #         layout.setContentsMargins(0, 0, 0, 0)
# # #         layout.setSpacing(0)

# # #         # ── Title bar ──
# # #         title_bar = QWidget()
# # #         title_bar.setStyleSheet(
# # #             f"background-color: {PANEL}; "
# # #             f"border-top-left-radius: {_CARD_RADIUS}px; "
# # #             f"border-top-right-radius: {_CARD_RADIUS}px; "
# # #             f"border-bottom: 1px solid rgba(255, 255, 255, 0.04);"
# # #         )
# # #         tb_layout = QHBoxLayout(title_bar)
# # #         tb_layout.setContentsMargins(16, 10, 16, 10)

# # #         for dot_color in (ERROR, SECONDARY, TERTIARY):
# # #             dot = QLabel()
# # #             dot.setFixedSize(10, 10)
# # #             dot.setStyleSheet(
# # #                 f"background-color: {dot_color}; border-radius: 5px;"
# # #             )
# # #             tb_layout.addWidget(dot)
# # #             tb_layout.addSpacing(6)

# # #         tb_layout.addSpacing(12)
# # #         term_icon_label = QLabel()
# # #         term_icon_label.setPixmap(
# # #             qta.icon("fa6s.terminal", color=TEXT_MUTED).pixmap(12, 12)
# # #         )
# # #         tb_layout.addWidget(term_icon_label)
# # #         tb_layout.addSpacing(6)
# # #         self.term_label = QLabel(self._DEFAULT_TITLE)
# # #         self.term_label.setStyleSheet(
# # #             f"color: {TEXT_MUTED}; font-size: 8.5pt;"
# # #         )
# # #         tb_layout.addWidget(self.term_label)
# # #         tb_layout.addStretch()

# # #         layout.addWidget(title_bar)

# # #         # ── Log view ──
# # #         self.log = QTextEdit()
# # #         self.log.setReadOnly(True)
# # #         self.log.setMinimumHeight(180)
# # #         self.log.setFont(QFont("Consolas", 10))
# # #         self.log.setPlaceholderText("")
# # #         self.log.setStyleSheet(
# # #             f"QTextEdit {{ background-color: transparent; "
# # #             f"color: {TEXT_SECONDARY}; border: none; padding: 16px 20px; "
# # #             f"selection-background-color: rgba(99, 102, 241, 0.3); }}"
# # #         )
# # #         layout.addWidget(self.log, stretch=1)

# # #         # ── Processing animation ──
# # #         self._dots_timer = QTimer(self)
# # #         self._dots_timer.setInterval(_ANIM_SLOW)
# # #         self._dots_timer.timeout.connect(self._tick_dots)
# # #         self._dots_count = 0

# # #         self._title_bar = title_bar
# # #         self._pulse_opacity = QGraphicsOpacityEffect(title_bar)
# # #         title_bar.setGraphicsEffect(self._pulse_opacity)
# # #         self._pulse_opacity.setOpacity(1.0)

# # #         self._pulse_anim = QPropertyAnimation(self._pulse_opacity, b"opacity")
# # #         self._pulse_anim.setDuration(900)
# # #         self._pulse_anim.setStartValue(1.0)
# # #         self._pulse_anim.setEndValue(0.55)
# # #         self._pulse_anim.setEasingCurve(QEasingCurve.InOutSine)
# # #         self._pulse_anim.setLoopCount(-1)

# # #     def start_processing(self):
# # #         self._dots_count = 0
# # #         self.term_label.setText("PROCESSING.")
# # #         self._dots_timer.start()
# # #         self._pulse_opacity.setOpacity(1.0)
# # #         self._pulse_anim.stop()
# # #         self._pulse_anim.setDirection(QPropertyAnimation.Forward)
# # #         self._pulse_anim.start()

# # #     def stop_processing(self):
# # #         self._dots_timer.stop()
# # #         self._pulse_anim.stop()
# # #         self._pulse_opacity.setOpacity(1.0)
# # #         self.term_label.setText(self._DEFAULT_TITLE)

# # #     def _tick_dots(self):
# # #         self._dots_count = (self._dots_count % 3) + 1
# # #         self.term_label.setText("PROCESSING" + "." * self._dots_count)

# # #     def append(self, html):
# # #         self.log.append(html)

# # #     def clear(self):
# # #         self.log.clear()

# # #     def set_placeholder(self, html):
# # #         self.log.setHtml(html)


# # # class DropZone(QFrame):
# # #     """Clickable + drag-and-drop file picker card with shadow elevation."""

# # #     def __init__(self, icon, icon_color, title, hint, extensions, badge_text, on_pick):
# # #         super().__init__()
# # #         self.extensions = extensions
# # #         self.on_pick = on_pick
# # #         self.path = None

# # #         self.setAcceptDrops(True)
# # #         self.setCursor(Qt.PointingHandCursor)
# # #         self.setMinimumHeight(200)
# # #         self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
# # #         self.setStyleSheet(_dropzone_stylesheet())
# # #         self.setToolTip(
# # #             f"Select or drag a file ({', '.join(extensions)})"
# # #         )

# # #         self._shadow = _make_shadow(self)
# # #         self.setGraphicsEffect(self._shadow)

# # #         self._shadow_anim = QPropertyAnimation(self._shadow, b"blurRadius")
# # #         self._shadow_anim.setDuration(_ANIM_NORMAL)
# # #         self._shadow_anim.setEasingCurve(QEasingCurve.OutCubic)

# # #         self._shadow_color_anim = QPropertyAnimation(self._shadow, b"color")
# # #         self._shadow_color_anim.setDuration(_ANIM_NORMAL)
# # #         self._shadow_color_anim.setEasingCurve(QEasingCurve.OutCubic)

# # #         layout = QVBoxLayout(self)
# # #         layout.setContentsMargins(
# # #             _CARD_PADDING_H, _CARD_PADDING_V, _CARD_PADDING_H, _CARD_PADDING_V
# # #         )
# # #         layout.setSpacing(10)
# # #         layout.setAlignment(Qt.AlignCenter)

# # #         icon_label = QLabel()
# # #         icon_label.setAlignment(Qt.AlignCenter)
# # #         icon_label.setFixedSize(_ICON_SIZE, _ICON_SIZE)
# # #         icon_label.setStyleSheet(
# # #             f"background-color: rgba(255, 255, 255, 0.03); "
# # #             f"border-radius: {_ICON_PADDING}px;"
# # #         )
# # #         icon_pixmap = qta.icon(icon, color=icon_color).pixmap(32, 32)
# # #         icon_label.setPixmap(icon_pixmap)

# # #         icon_row = QHBoxLayout()
# # #         icon_row.addStretch()
# # #         icon_row.addWidget(icon_label)
# # #         icon_row.addStretch()

# # #         self.title_label = QLabel(title)
# # #         self.title_label.setAlignment(Qt.AlignCenter)
# # #         self.title_label.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
# # #         self.title_label.setStyleSheet(
# # #             f"color: {TEXT_PRIMARY}; background: transparent;"
# # #         )

# # #         self.badge_label = QLabel(badge_text)
# # #         self.badge_label.setAlignment(Qt.AlignCenter)
# # #         self.badge_label.setWordWrap(True)
# # #         self.badge_label.setStyleSheet(
# # #             f"color: {icon_color}; background-color: rgba(255, 255, 255, 0.04); "
# # #             f"border: 1px solid rgba(255, 255, 255, 0.06); "
# # #             f"border-radius: {_BADGE_RADIUS}px; "
# # #             f"padding: 4px 14px; font-size: 10pt;"
# # #         )

# # #         self.hint_label = QLabel(hint)
# # #         self.hint_label.setAlignment(Qt.AlignCenter)
# # #         self.hint_label.setWordWrap(True)
# # #         self.hint_label.setStyleSheet(
# # #             f"color: {TEXT_MUTED}; font-size: 10pt; background: transparent;"
# # #         )

# # #         self.filename_label = QLabel("")
# # #         self.filename_label.setAlignment(Qt.AlignCenter)
# # #         self.filename_label.setWordWrap(True)
# # #         self.filename_label.setStyleSheet(
# # #             f"color: {SUCCESS}; font-size: 10pt; background: transparent;"
# # #         )
# # #         self.filename_label.hide()

# # #         layout.addLayout(icon_row)
# # #         layout.addWidget(self.title_label)
# # #         layout.addWidget(self.badge_label)
# # #         layout.addWidget(self.hint_label)
# # #         layout.addWidget(self.filename_label)

# # #         self._default_badge = badge_text
# # #         self._icon_color = icon_color

# # #     # ── Shadow animation ──────────────────────────────────────────────

# # #     def _animate_shadow(self, target_blur, target_color=None, duration=_ANIM_NORMAL):
# # #         self._shadow_anim.stop()
# # #         self._shadow_anim.setDuration(duration)
# # #         self._shadow_anim.setStartValue(self._shadow.blurRadius())
# # #         self._shadow_anim.setEndValue(target_blur)
# # #         self._shadow_anim.start()

# # #         if target_color is not None:
# # #             self._shadow_color_anim.stop()
# # #             self._shadow_color_anim.setDuration(duration)
# # #             self._shadow_color_anim.setStartValue(self._shadow.color())
# # #             self._shadow_color_anim.setEndValue(target_color)
# # #             self._shadow_color_anim.start()

# # #     def _hover_shadow_color(self):
# # #         if self.path:
# # #             return QColor(52, 211, 153, 25)
# # #         return QColor(99, 102, 241, 40)

# # #     # ── Hover events ──────────────────────────────────────────────────

# # #     def enterEvent(self, event):
# # #         self._animate_shadow(
# # #             _SHADOW_HOVER,
# # #             target_color=self._hover_shadow_color(),
# # #             duration=_ANIM_FAST,
# # #         )
# # #         super().enterEvent(event)

# # #     def leaveEvent(self, event):
# # #         self._animate_shadow(
# # #             _SHADOW_DEFAULT,
# # #             target_color=QColor(0, 0, 0, 60),
# # #             duration=_ANIM_NORMAL,
# # #         )
# # #         super().leaveEvent(event)

# # #     # ── Existing logic (unchanged) ────────────────────────────────────

# # #     def _set_style(self, active: bool, filled: bool = False):
# # #         self.setStyleSheet(_dropzone_stylesheet(active=active, filled=filled))

# # #     def _valid_path(self, path: str) -> bool:
# # #         return path.lower().endswith(tuple(self.extensions))

# # #     def _apply_path(self, path: str):
# # #         self.path = path
# # #         self.badge_label.setText(os.path.basename(path))
# # #         self.filename_label.hide()
# # #         self.hint_label.setText("Click or drop again to replace")
# # #         self._set_style(active=False, filled=True)
# # #         self._animate_shadow(
# # #             _SHADOW_DEFAULT,
# # #             target_color=QColor(0, 0, 0, 60),
# # #         )
# # #         self.on_pick()

# # #     def mousePressEvent(self, event):
# # #         filt = " ".join(f"*{e}" for e in self.extensions)
# # #         path, _ = QFileDialog.getOpenFileName(
# # #             self, "Select file", "", f"Supported Files ({filt})"
# # #         )
# # #         if path and self._valid_path(path):
# # #             self._apply_path(path)

# # #     def dragEnterEvent(self, event: QDragEnterEvent):
# # #         if event.mimeData().hasUrls():
# # #             url = event.mimeData().urls()[0]
# # #             if self._valid_path(url.toLocalFile()):
# # #                 self._set_style(active=True)
# # #                 self._animate_shadow(
# # #                     _SHADOW_DRAG,
# # #                     target_color=QColor(99, 102, 241, 70),
# # #                     duration=_ANIM_FAST,
# # #                 )
# # #                 event.acceptProposedAction()
# # #                 return
# # #         event.ignore()

# # #     def dragLeaveEvent(self, event):
# # #         self._set_style(active=False, filled=bool(self.path))
# # #         self._animate_shadow(
# # #             _SHADOW_DEFAULT,
# # #             target_color=QColor(0, 0, 0, 60),
# # #             duration=_ANIM_NORMAL,
# # #         )

# # #     def dropEvent(self, event: QDropEvent):
# # #         url = event.mimeData().urls()[0]
# # #         path = url.toLocalFile()
# # #         if self._valid_path(path):
# # #             self._apply_path(path)
# # #             event.acceptProposedAction()
# # #         else:
# # #             self._set_style(active=False, filled=bool(self.path))
# # #             self._animate_shadow(
# # #                 _SHADOW_DEFAULT,
# # #                 target_color=QColor(0, 0, 0, 60),
# # #                 duration=_ANIM_NORMAL,
# # #             )
# # #             event.ignore()


# # # # ═══════════════════════════════════════════════════════════════════════════════
# # # # Background Worker (unchanged)
# # # # ═══════════════════════════════════════════════════════════════════════════════


# # # class WorkerThread(QThread):
# # #     finished_ok = Signal(dict)
# # #     finished_err = Signal(str)

# # #     def __init__(self, excel_path, data_path, skip_unrelated_sheets):
# # #         super().__init__()
# # #         self.excel_path = excel_path
# # #         self.data_path = data_path
# # #         self.skip_unrelated_sheets = skip_unrelated_sheets

# # #     def run(self):
# # #         try:
# # #             summary = fill_workbook(
# # #                 self.excel_path,
# # #                 self.data_path,
# # #                 skip_unrelated_sheets=self.skip_unrelated_sheets,
# # #             )
# # #             self.finished_ok.emit(summary)
# # #         except Exception:
# # #             self.finished_err.emit(traceback.format_exc())


# # # # ═══════════════════════════════════════════════════════════════════════════════
# # # # Main Window
# # # # ═══════════════════════════════════════════════════════════════════════════════


# # # class MainWindow(QMainWindow):
# # #     def __init__(self):
# # #         super().__init__()
# # #         self.setWindowTitle("Excel Auto-Fill Bot")
# # #         self.resize(860, 780)
# # #         self.setMinimumSize(700, 720)
# # #         self.setStyleSheet(f"QMainWindow {{ background-color: {BG}; }}")
# # #         if os.path.exists(LOGO_PATH):
# # #             self.setWindowIcon(QIcon(LOGO_PATH))

# # #         central = QWidget()
# # #         self.setCentralWidget(central)
# # #         root = QVBoxLayout(central)
# # #         root.setContentsMargins(0, 0, 0, 0)
# # #         root.setSpacing(0)

# # #         root.addWidget(self._build_header())

# # #         body = QWidget()
# # #         body_layout = QVBoxLayout(body)
# # #         body_layout.setContentsMargins(
# # #             _BODY_MARGIN_H, _BODY_MARGIN_V, _BODY_MARGIN_H, _BODY_MARGIN_V
# # #         )
# # #         body_layout.setSpacing(_SECTION_SPACING)

# # #         # ── Drop zones ──
# # #         picker_row = QHBoxLayout()
# # #         picker_row.setSpacing(16)
# # #         self.excel_zone = DropZone(
# # #             "fa6s.table", TERTIARY, "Excel File",
# # #             "Drag & drop .xlsx here\nor click to browse",
# # #             [".xlsx", ".xlsm"], "multisheet_workbook.xlsx",
# # #             self._update_run_state,
# # #         )
# # #         self.data_zone = DropZone(
# # #             "fa6s.file-code", SECONDARY, "Data File (JSON / XML)",
# # #             "Drag & drop .json / .xml here\nor click to browse",
# # #             [".json", ".xml"], "response.json",
# # #             self._update_run_state,
# # #         )
# # #         picker_row.addWidget(self.excel_zone)
# # #         picker_row.addWidget(self.data_zone)
# # #         body_layout.addLayout(picker_row)

# # #         # ── Options ──
# # #         self.fill_all_checkbox = QCheckBox(
# # #             "Also fill sheets that don't match this response at all "
# # #             "(fills every cell with N/A)"
# # #         )
# # #         self.fill_all_checkbox.setChecked(False)
# # #         self.fill_all_checkbox.setToolTip(
# # #             "When enabled, sheets with no matching headers "
# # #             "will still receive a row filled entirely with N/A"
# # #         )
# # #         self.fill_all_checkbox.setStyleSheet(f"""
# # #             QCheckBox {{
# # #                 color: {TEXT_SECONDARY};
# # #                 font-size: 10pt;
# # #                 spacing: 10px;
# # #                 padding-left: 4px;
# # #             }}
# # #             QCheckBox::indicator {{
# # #                 width: {_INDICATOR_SIZE}px;
# # #                 height: {_INDICATOR_SIZE}px;
# # #                 border-radius: 5px;
# # #                 border: 1.5px solid rgba(255, 255, 255, 0.15);
# # #                 background-color: transparent;
# # #             }}
# # #             QCheckBox::indicator:checked {{
# # #                 background-color: {ACCENT_CONTAINER};
# # #                 border: 1.5px solid {ACCENT_CONTAINER};
# # #             }}
# # #             QCheckBox::indicator:hover {{
# # #                 border-color: {ACCENT_HOVER};
# # #             }}
# # #         """)
# # #         body_layout.addWidget(self.fill_all_checkbox)

# # #         # ── Run button ──
# # #         self.run_btn = PrimaryButton("\u25b6  Run \u2014 Fill Excel File")
# # #         self.run_btn.clicked.connect(self._run)
# # #         body_layout.addWidget(self.run_btn)

# # #         # ── Status pill ──
# # #         self.status_label = StatusPill()
# # #         body_layout.addWidget(self.status_label)

# # #         # ── Terminal ──
# # #         self.terminal = TerminalWidget()
# # #         self.log = self.terminal.log
# # #         body_layout.addWidget(self.terminal, stretch=1)

# # #         root.addWidget(body, stretch=1)

# # #         self.worker = None
# # #         self._log_placeholder()

# # #     # ── Header ────────────────────────────────────────────────────────────

# # #     def _build_header(self):
# # #         header = QWidget()
# # #         header.setFixedHeight(72)
# # #         header.setStyleSheet(
# # #             f"background-color: {SURFACE}; "
# # #             f"border-bottom: 1px solid {BORDER};"
# # #         )
# # #         layout = QVBoxLayout(header)
# # #         layout.setContentsMargins(
# # #             _BODY_MARGIN_H, 0, _BODY_MARGIN_H, 0
# # #         )
# # #         layout.setSpacing(0)

# # #         row = QHBoxLayout()
# # #         row.setSpacing(12)

# # #         badge = QLabel()
# # #         badge.setFixedSize(40, 40)
# # #         badge.setAlignment(Qt.AlignCenter)
# # #         badge.setStyleSheet(
# # #             f"background-color: {ACCENT_SOFT}; border-radius: 10px;"
# # #         )
# # #         if os.path.exists(LOGO_PATH):
# # #             badge.setPixmap(
# # #                 QPixmap(LOGO_PATH).scaled(
# # #                     40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation
# # #                 )
# # #             )
# # #         else:
# # #             badge_icon = qta.icon("fa6s.bolt", color=ACCENT_CONTAINER)
# # #             badge.setPixmap(badge_icon.pixmap(20, 20))

# # #         title_box = QVBoxLayout()
# # #         title_box.setSpacing(1)
# # #         heading = QLabel("Excel Auto-Fill Bot")
# # #         heading.setFont(QFont("Segoe UI", 17, QFont.DemiBold))
# # #         heading.setStyleSheet(f"color: {TEXT_PRIMARY}; background: transparent;")
# # #         title_box.addWidget(heading)

# # #         subtitle = QLabel("Automated spreadsheet filling")
# # #         subtitle.setFont(QFont("Segoe UI", 9))
# # #         subtitle.setStyleSheet(f"color: {TEXT_MUTED}; background: transparent;")
# # #         title_box.addWidget(subtitle)

# # #         row.addWidget(badge)
# # #         row.addSpacing(4)
# # #         row.addLayout(title_box)
# # #         row.addStretch()
# # #         layout.addLayout(row)

# # #         return header

# # #     # ── Log helpers ───────────────────────────────────────────────────────

# # #     def _log_placeholder(self):
# # #         self.terminal.set_placeholder(
# # #             f"<span style='color:{TEXT_MUTED};'>Waiting for files\u2026 select or "
# # #             f"drop an Excel file and a JSON/XML data file above, then hit Run.</span>"
# # #         )

# # #     def _append_log(self, html):
# # #         self.terminal.append(html)

# # #     # ── State management ──────────────────────────────────────────────────

# # #     def _update_run_state(self):
# # #         self.run_btn.setEnabled(bool(self.excel_zone.path and self.data_zone.path))

# # #     def _set_status(self, text, kind="info"):
# # #         self.status_label.set_status(text, kind)

# # #     # ── Run / results ─────────────────────────────────────────────────────

# # #     def _run(self):
# # #         self.run_btn.setEnabled(False)
# # #         self.run_btn.setText("Processing\u2026")
# # #         self.status_label.hide()
# # #         self.terminal.clear()
# # #         self.terminal.start_processing()
# # #         self._append_log(
# # #             f"<b style='color:{TEXT_PRIMARY}'>Excel:</b> {self.excel_zone.path}"
# # #         )
# # #         self._append_log(
# # #             f"<b style='color:{TEXT_PRIMARY}'>Data:</b> {self.data_zone.path}"
# # #         )
# # #         self._append_log(
# # #             f"<span style='color:{TEXT_MUTED}'>Reading sheets, detecting "
# # #             f"headers, matching data\u2026</span><br>"
# # #         )

# # #         skip_unrelated = not self.fill_all_checkbox.isChecked()
# # #         self.worker = WorkerThread(
# # #             self.excel_zone.path, self.data_zone.path, skip_unrelated
# # #         )
# # #         self.worker.finished_ok.connect(self._on_success)
# # #         self.worker.finished_err.connect(self._on_error)
# # #         self.worker.start()

# # #     def _on_success(self, summary):
# # #         self.terminal.stop_processing()
# # #         self.run_btn.setEnabled(True)
# # #         self.run_btn.setText("\u25b6  Run \u2014 Fill Excel File")

# # #         self._append_log(
# # #             f"<span style='color:{TEXT_PRIMARY}'>Records found in data file: "
# # #             f"<b>{summary['records_processed']}</b></span><br>"
# # #         )
# # #         for sheet in summary["sheets"]:
# # #             self._append_log(
# # #                 f"<b style='color:{SECONDARY}'>\u25a0 {sheet['sheet']}</b>"
# # #             )
# # #             if sheet["skipped"]:
# # #                 self._append_log(
# # #                     f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;No matching "
# # #                     f"headers \u2014 sheet left untouched.</span><br>"
# # #                 )
# # #                 continue
# # #             self._append_log(
# # #                 f"<span style='color:{TERTIARY}'>&nbsp;&nbsp;Rows filled: "
# # #                 f"{sheet['rows_filled']} (new: {sheet['rows_added']}, "
# # #                 f"replaced duplicates: {sheet['rows_replaced']})</span>"
# # #             )
# # #             if sheet["matched_headers"]:
# # #                 self._append_log(
# # #                     f"<span style='color:{TEXT_SECONDARY}'>"
# # #                     f"&nbsp;&nbsp;Matched columns:</span>"
# # #                 )
# # #                 for header, key in sheet["matched_headers"].items():
# # #                     self._append_log(
# # #                         f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;&nbsp;&nbsp;"
# # #                         f"'{header}' \u2190 '{key}'</span>"
# # #                     )
# # #             if sheet["unmatched_headers"]:
# # #                 self._append_log(
# # #                     f"<span style='color:{ERROR}'>&nbsp;&nbsp;Filled with N/A "
# # #                     f"(no matching data): "
# # #                     f"{', '.join(sheet['unmatched_headers'])}</span>"
# # #                 )
# # #             self._append_log("")

# # #         self._append_log(
# # #             f"<b style='color:{TERTIARY}'>\u2713 Done \u2014 Excel file updated "
# # #             f"in place.</b>"
# # #         )
# # #         self._set_status(
# # #             "\u2713 Excel file filled and saved successfully.", kind="success"
# # #         )

# # #     def _on_error(self, err_text):
# # #         self.terminal.stop_processing()
# # #         self.run_btn.setEnabled(True)
# # #         self.run_btn.setText("\u25b6  Run \u2014 Fill Excel File")
# # #         self._append_log(
# # #             f"<span style='color:{ERROR}'><b>ERROR</b></span>"
# # #         )
# # #         self._append_log(f"<span style='color:{ERROR}'>{err_text}</span>")
# # #         self._set_status(
# # #             "\u2717 Something went wrong \u2014 see log for details.", kind="error"
# # #         )


# # # # ═══════════════════════════════════════════════════════════════════════════════
# # # # Entry Point
# # # # ═══════════════════════════════════════════════════════════════════════════════


# # # def main():
# # #     app = QApplication(sys.argv)
# # #     app.setStyle("Fusion")
# # #     app.setFont(QFont("Segoe UI", 10))
# # #     app.setStyleSheet(f"""
# # #         QToolTip {{
# # #             background-color: {PANEL_LIGHT};
# # #             color: {TEXT_PRIMARY};
# # #             border: 1px solid rgba(255, 255, 255, 0.10);
# # #             border-radius: 6px;
# # #             padding: 6px 10px;
# # #             font-size: 10pt;
# # #         }}
# # #         QScrollBar:vertical {{
# # #             background: transparent;
# # #             width: 8px;
# # #             margin: 0;
# # #         }}
# # #         QScrollBar::handle:vertical {{
# # #             background: rgba(255, 255, 255, 0.12);
# # #             min-height: 30px;
# # #             border-radius: 4px;
# # #         }}
# # #         QScrollBar::handle:vertical:hover {{
# # #             background: rgba(255, 255, 255, 0.20);
# # #         }}
# # #         QScrollBar::add-line:vertical,
# # #         QScrollBar::sub-line:vertical {{
# # #             height: 0;
# # #         }}
# # #         QScrollBar::add-page:vertical,
# # #         QScrollBar::sub-page:vertical {{
# # #             background: transparent;
# # #         }}
# # #         QScrollBar:horizontal {{
# # #             background: transparent;
# # #             height: 8px;
# # #             margin: 0;
# # #         }}
# # #         QScrollBar::handle:horizontal {{
# # #             background: rgba(255, 255, 255, 0.12);
# # #             min-width: 30px;
# # #             border-radius: 4px;
# # #         }}
# # #         QScrollBar::handle:horizontal:hover {{
# # #             background: rgba(255, 255, 255, 0.20);
# # #         }}
# # #         QScrollBar::add-line:horizontal,
# # #         QScrollBar::sub-line:horizontal {{
# # #             width: 0;
# # #         }}
# # #         QScrollBar::add-page:horizontal,
# # #         QScrollBar::sub-page:horizontal {{
# # #             background: transparent;
# # #         }}
# # #     """)
# # #     if os.path.exists(LOGO_PATH):
# # #         app.setWindowIcon(QIcon(LOGO_PATH))
# # #     win = MainWindow()
# # #     win.show()
# # #     sys.exit(app.exec())


# # # if __name__ == "__main__":
# # #     main()
















# """
# Excel Auto-Fill Bot - Desktop App

# Pick (or drag & drop) an Excel file (multi-sheet, dynamic headers - nothing
# hardcoded) and a JSON or XML data file (e.g. an Aadhaar-style API response).
# The app fuzzy-matches the data's keys against each sheet's headers, fills
# matching values into a brand new row of every sheet, writes "N/A" for
# anything with no matching data, and overwrites the same Excel file.

# - Har record (chahe duplicate hi kyu na ho) apni alag row leta hai - kuch bhi
#   skip/merge nahi hota.
# - By default koi bhi sheet skip nahi hoti - jis sheet mein match nahi milta
#   wahan bhi ek row jaake N/A se bhar jaati hai. "Sirf unrelated sheets skip
#   karo" checkbox se old (strict) behaviour wapas mil jata hai.
# """

# import os
# import sys
# import traceback

# from PySide6.QtCore import Qt, QThread, Signal
# from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent, QColor, QIcon, QPixmap
# from PySide6.QtWidgets import (
#     QApplication,
#     QCheckBox,
#     QFileDialog,
#     QFrame,
#     QGraphicsDropShadowEffect,
#     QHBoxLayout,
#     QLabel,
#     QMainWindow,
#     QPushButton,
#     QSizePolicy,
#     QTextEdit,
#     QVBoxLayout,
#     QWidget,
# )

# from filler_core import fill_workbook

# # ---------------------------------------------------------------------------
# # Palette (matches the Excel Auto-Fill Bot design spec)
# # ---------------------------------------------------------------------------
# BG = "#0b1326"
# SURFACE = "#0b1326"
# PANEL = "#171f33"
# PANEL_LIGHT = "#1e2338"
# SURFACE_HIGH = "#222a3d"
# SURFACE_HIGHEST = "#2d3449"
# BORDER = "#464554"
# BORDER_ACTIVE = "#8083ff"
# ACCENT = "#c0c1ff"                 # primary
# ACCENT_CONTAINER = "#8083ff"       # primary-container
# ACCENT_HOVER = "#6f72e8"
# ACCENT_SOFT = "#2a2d55"
# SECONDARY = "#89ceff"
# SECONDARY_CONTAINER = "#00a2e6"
# TERTIARY = "#4edea3"
# TERTIARY_CONTAINER = "#00885d"
# TEXT_PRIMARY = "#dae2fd"
# TEXT_SECONDARY = "#c7c4d7"
# TEXT_MUTED = "#908fa0"
# SUCCESS = "#4edea3"
# ERROR = "#ffb4ab"

# ASSET_DIR = os.path.dirname(os.path.abspath(__file__))
# LOGO_PATH = os.path.join(ASSET_DIR, "logo.png")
# CHECK_ICON_PATH = os.path.join(ASSET_DIR, "check_icon.svg").replace("\\", "/")


# def _checkbox_style() -> str:
#     """
#     Clean, modern checkbox: a rounded square that's a soft outline when
#     unchecked and fills solid with a crisp checkmark when checked, with a
#     subtle hover highlight. Qt's stylesheet engine takes over the indicator
#     entirely once any indicator property is set, so the checkmark glyph is
#     supplied explicitly via CHECK_ICON_PATH rather than left to the OS theme.
#     """
#     return f"""
#         QCheckBox {{
#             spacing: 12px;
#             color: {TEXT_SECONDARY};
#             font-size: 10.5pt;
#             background: transparent;
#             border: none;
#             padding: 2px 0;
#         }}
#         QCheckBox::indicator {{
#             width: 19px;
#             height: 19px;
#             border-radius: 6px;
#             border: 1.5px solid {BORDER};
#             background-color: {SURFACE_HIGH};
#         }}
#         QCheckBox::indicator:hover {{
#             border-color: {BORDER_ACTIVE};
#             background-color: {ACCENT_SOFT};
#         }}
#         QCheckBox::indicator:checked {{
#             border-color: {ACCENT_CONTAINER};
#             background-color: {ACCENT_CONTAINER};
#             image: url({CHECK_ICON_PATH});
#         }}
#         QCheckBox::indicator:checked:hover {{
#             border-color: {ACCENT_HOVER};
#             background-color: {ACCENT_HOVER};
#         }}
#     """


# class DropZone(QFrame):
#     """Clickable + drag-and-drop file picker box (bento-grid card)."""

#     def __init__(self, icon, icon_color, title, hint, extensions, badge_text, on_pick):
#         super().__init__()
#         self.extensions = extensions
#         self.on_pick = on_pick
#         self.path = None

#         self.setAcceptDrops(True)
#         self.setCursor(Qt.PointingHandCursor)
#         self.setMinimumHeight(170)
#         self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
#         self._set_style(active=False)

#         shadow = QGraphicsDropShadowEffect(self)
#         shadow.setBlurRadius(28)
#         shadow.setOffset(0, 8)
#         shadow.setColor(QColor(0, 0, 0, 110))
#         self.setGraphicsEffect(shadow)

#         layout = QVBoxLayout(self)
#         layout.setContentsMargins(20, 24, 20, 24)
#         layout.setSpacing(8)
#         layout.setAlignment(Qt.AlignCenter)

#         icon_wrap = QLabel(icon)
#         icon_wrap.setAlignment(Qt.AlignCenter)
#         icon_wrap.setFixedSize(64, 64)
#         icon_wrap.setStyleSheet(
#             f"font-size: 26px; background-color: {SURFACE_HIGH}; border: none; border-radius: 18px; color: {icon_color};"
#         )

#         icon_row = QHBoxLayout()
#         icon_row.addStretch()
#         icon_row.addWidget(icon_wrap)
#         icon_row.addStretch()

#         self.title_label = QLabel(title)
#         self.title_label.setAlignment(Qt.AlignCenter)
#         self.title_label.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
#         self.title_label.setStyleSheet(f"""
# QLabel {{
#     color: {TEXT_PRIMARY};
#     background: transparent;
#     border: none;
# }}
# """)

#         self.badge_label = QLabel(badge_text)
#         self.badge_label.setAlignment(Qt.AlignCenter)
#         self.badge_label.setWordWrap(True)

#         self.hint_label = QLabel(hint)
#         self.hint_label.setAlignment(Qt.AlignCenter)
#         self.hint_label.setWordWrap(True)
#         self.hint_label.setStyleSheet(f"""
# QLabel {{
#  color: {TEXT_MUTED};
#  background: transparent;
#  border: none;
#  font-size:10pt;
# }}
# """)

#         layout.addLayout(icon_row)
#         layout.addWidget(self.title_label)
#         layout.addWidget(self.badge_label)
#         layout.addWidget(self.hint_label)

#         self._default_badge = badge_text
#         self._icon_color = icon_color
#         self._style_badge(filled=False)

#     def _style_badge(self, filled: bool):
#         if filled:
#             self.badge_label.setStyleSheet(f"""
# QLabel {{
#     color: {self._icon_color};
#     background-color: rgba(255,255,255,0.06);
#     border: none;
#     border-radius: 999px;
#     padding: 4px 14px;
#     font-size: 10.5pt;
#     font-weight: 600;
# }}
# """)
#         else:
#             self.badge_label.setStyleSheet(f"""
# QLabel {{
#     color: {TEXT_MUTED};
#     background-color: rgba(255,255,255,0.03);
#     border: 1px dashed {BORDER};
#     border-radius: 999px;
#     padding: 4px 14px;
#     font-size: 10pt;
#     font-style: italic;
# }}
# """)

#     def _set_style(self, active: bool, filled: bool = False):
#         if active:
#             border, bg = BORDER_ACTIVE, ACCENT_SOFT
#         elif filled:
#             border, bg = TERTIARY, PANEL_LIGHT
#         else:
#             border, bg = BORDER, PANEL
#         self.setStyleSheet(f"""
# DropZone {{
#     background-color: {bg};
#     border: 2px solid {border};
#     border-radius: 18px;
# }}
# QLabel {{
#     border:none;
#     background:transparent;
# }}
# """)

#     def _valid_path(self, path: str) -> bool:
#         return path.lower().endswith(tuple(self.extensions))

#     def _apply_path(self, path: str):
#         self.path = path
#         self.badge_label.setText(os.path.basename(path))
#         self._style_badge(filled=True)
#         self.hint_label.setText("Click or drop again to replace")
#         self._set_style(active=False, filled=True)
#         self.on_pick()

#     def mousePressEvent(self, event):
#         filt = " ".join(f"*{e}" for e in self.extensions)
#         path, _ = QFileDialog.getOpenFileName(self, "Select file", "", f"Supported Files ({filt})")
#         if path and self._valid_path(path):
#             self._apply_path(path)

#     def dragEnterEvent(self, event: QDragEnterEvent):
#         if event.mimeData().hasUrls():
#             url = event.mimeData().urls()[0]
#             if self._valid_path(url.toLocalFile()):
#                 self._set_style(active=True)
#                 event.acceptProposedAction()
#                 return
#         event.ignore()

#     def dragLeaveEvent(self, event):
#         self._set_style(active=False, filled=bool(self.path))

#     def dropEvent(self, event: QDropEvent):
#         url = event.mimeData().urls()[0]
#         path = url.toLocalFile()
#         if self._valid_path(path):
#             self._apply_path(path)
#             event.acceptProposedAction()
#         else:
#             self._set_style(active=False, filled=bool(self.path))
#             event.ignore()


# class WorkerThread(QThread):
#     finished_ok = Signal(dict)
#     finished_err = Signal(str)

#     def __init__(self, excel_path, data_path, skip_unrelated_sheets, save_path=None):
#         super().__init__()
#         self.excel_path = excel_path
#         self.data_path = data_path
#         self.skip_unrelated_sheets = skip_unrelated_sheets
#         self.save_path = save_path

#     def run(self):
#         try:
#             summary = fill_workbook(
#                 self.excel_path,
#                 self.data_path,
#                 save_path=self.save_path,
#                 skip_unrelated_sheets=self.skip_unrelated_sheets,
#             )
#             self.finished_ok.emit(summary)
#         except PermissionError as e:
#             self.finished_err.emit(str(e))
#         except Exception:
#             self.finished_err.emit(traceback.format_exc())


# class MainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("Excel Auto-Fill Bot")
#         self.resize(820, 760)
#         self.setMinimumSize(700, 620)
#         self.setStyleSheet(f"QMainWindow {{ background-color: {BG}; }}")
#         if os.path.exists(LOGO_PATH):
#             self.setWindowIcon(QIcon(LOGO_PATH))

#         central = QWidget()
#         self.setCentralWidget(central)
#         root = QVBoxLayout(central)
#         root.setContentsMargins(0, 0, 0, 0)
#         root.setSpacing(0)

#         # ---------- Header (AppBar) ----------
#         header = QWidget()
#         header.setStyleSheet(f"background-color: {SURFACE}; border-bottom: 1px solid {BORDER};")
#         header_layout = QVBoxLayout(header)
#         header_layout.setContentsMargins(32, 26, 32, 22)
#         header_layout.setSpacing(4)

#         header_row = QHBoxLayout()
#         badge = QLabel()
#         badge.setFixedSize(48, 48)
#         if os.path.exists(LOGO_PATH):
#             badge.setPixmap(
#                 QPixmap(LOGO_PATH).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
#             )
#         else:
#             badge.setText("⚡")
#             badge.setStyleSheet(
#                 f"background-color: {ACCENT_SOFT}; border-radius: 12px; font-size: 20px; color: {ACCENT};"
#             )
#         badge.setAlignment(Qt.AlignCenter)

#         title_box = QVBoxLayout()
#         title_box.setSpacing(2)
#         heading = QLabel("Excel Auto-Fill Bot")
#         heading.setFont(QFont("Segoe UI", 21, QFont.Bold))
#         heading.setStyleSheet(f"color: {TEXT_PRIMARY};")
#         title_box.addWidget(heading)

#         subtitle = QLabel("Automated spreadsheet filling from JSON / XML")
#         subtitle.setFont(QFont("Segoe UI", 9))
#         subtitle.setStyleSheet(f"color: {TEXT_MUTED}; background: transparent; border: none;")
#         title_box.addWidget(subtitle)

#         header_row.addWidget(badge)
#         header_row.addSpacing(10)
#         header_row.addLayout(title_box)
#         header_row.addStretch()
#         header_layout.addLayout(header_row)

#         root.addWidget(header)

#         # ---------- Body ----------
#         body = QWidget()
#         body_layout = QVBoxLayout(body)
#         body_layout.setContentsMargins(32, 28, 32, 28)
#         body_layout.setSpacing(20)

#         # Drop zones (bento grid)
#         picker_row = QHBoxLayout()
#         picker_row.setSpacing(16)
#         self.excel_zone = DropZone(
#             "📊", TERTIARY, "Excel File",
#             "Drag & drop .xlsx here\nor click to browse",
#             [".xlsx", ".xlsm"], "No file selected", self._update_run_state,
#         )
#         self.data_zone = DropZone(
#             "🗂", SECONDARY, "Data File (JSON / XML)",
#             "Drag & drop .json / .xml here\nor click to browse",
#             [".json", ".xml"], "No file selected", self._update_run_state,
#         )
#         picker_row.addWidget(self.excel_zone)
#         picker_row.addWidget(self.data_zone)
#         body_layout.addLayout(picker_row)

#         # ---------- Options card ----------
#         options_card = QFrame()
#         options_card.setStyleSheet(
#             f"QFrame {{ background-color: {PANEL}; border: 1px solid {BORDER}; border-radius: 14px; }}"
#         )
#         options_layout = QVBoxLayout(options_card)
#         options_layout.setContentsMargins(20, 16, 20, 16)
#         options_layout.setSpacing(12)

#         options_title = QLabel("OPTIONS")
#         options_title.setStyleSheet(
#             f"color: {TEXT_MUTED}; font-size: 8.5pt; letter-spacing: 1px; "
#             f"background: transparent; border: none;"
#         )
#         options_layout.addWidget(options_title)

#         def _add_option(checkbox: QCheckBox, description: str):
#             checkbox.setCursor(Qt.PointingHandCursor)
#             checkbox.setStyleSheet(_checkbox_style())
#             row = QVBoxLayout()
#             row.setSpacing(2)
#             row.setContentsMargins(0, 0, 0, 0)
#             row.addWidget(checkbox)
#             desc = QLabel(description)
#             desc.setWordWrap(True)
#             desc.setStyleSheet(
#                 f"color: {TEXT_MUTED}; font-size: 9pt; background: transparent; "
#                 f"border: none; margin-left: 31px;"
#             )
#             row.addWidget(desc)
#             options_layout.addLayout(row)

#         self.fill_all_checkbox = QCheckBox("Fill every sheet, even unrelated ones")
#         self.fill_all_checkbox.setChecked(False)
#         _add_option(
#             self.fill_all_checkbox,
#             "Sheets with no matching columns still get a new row, filled entirely with N/A.",
#         )

#         divider = QFrame()
#         divider.setFixedHeight(1)
#         divider.setStyleSheet(f"background-color: {BORDER}; border: none;")
#         options_layout.addWidget(divider)

#         self.save_as_checkbox = QCheckBox("Save as a new Excel file")
#         self.save_as_checkbox.setChecked(False)
#         _add_option(
#             self.save_as_checkbox,
#             "Keeps the original workbook untouched and writes the result to a file you choose.",
#         )

#         body_layout.addWidget(options_card)

#         # Run button
#         self.run_btn = QPushButton("▶  Run — Fill Excel File")
#         self.run_btn.setEnabled(False)
#         self.run_btn.setCursor(Qt.PointingHandCursor)
#         self.run_btn.setMinimumHeight(56)
#         self.run_btn.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
#         self.run_btn.setStyleSheet(
#             f"""
#             QPushButton {{
#                 background-color: {ACCENT_CONTAINER};
#                 color: #07006c;
#                 border-radius: 12px;
#                 border: none;
#             }}
#             QPushButton:disabled {{
#                 background-color: {PANEL_LIGHT};
#                 color: {TEXT_MUTED};
#             }}
#             QPushButton:hover:!disabled {{
#                 background-color: {ACCENT_HOVER};
#                 color: #ffffff;
#             }}
#             QPushButton:pressed:!disabled {{
#                 background-color: #4f5bd1;
#                 color: #ffffff;
#             }}
#             """
#         )
#         self.run_btn.clicked.connect(self._run)
#         body_layout.addWidget(self.run_btn)

#         # Status pill
#         self.status_label = QLabel("")
#         self.status_label.setFont(QFont("Segoe UI", 10, QFont.Medium))
#         self.status_label.setAlignment(Qt.AlignCenter)
#         self.status_label.setMinimumHeight(40)
#         self.status_label.hide()
#         body_layout.addWidget(self.status_label)

#         # Activity log (terminal style)
#         log_card = QFrame()
#         log_card.setStyleSheet(
#             f"QFrame {{ background-color: {PANEL}; border: 1px solid {BORDER}; border-radius: 14px; }}"
#         )
#         log_card_layout = QVBoxLayout(log_card)
#         log_card_layout.setContentsMargins(0, 0, 0, 0)
#         log_card_layout.setSpacing(0)

#         term_bar = QWidget()
#         term_bar.setStyleSheet(
#             f"background-color: {SURFACE_HIGH}; border-top-left-radius: 14px; border-top-right-radius: 14px;"
#         )
#         term_bar_layout = QHBoxLayout(term_bar)
#         term_bar_layout.setContentsMargins(14, 8, 14, 8)
#         for dot_color in (ERROR, SECONDARY, TERTIARY):
#             dot = QLabel()
#             dot.setFixedSize(10, 10)
#             dot.setStyleSheet(f"background-color: {dot_color}; border-radius: 5px;")
#             term_bar_layout.addWidget(dot)
#             term_bar_layout.addSpacing(4)
#         term_bar_layout.addSpacing(10)
#         term_label = QLabel("TERMINAL — PROCESS_MAPPING")
#         term_label.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 8.5pt; letter-spacing: 1px;")
#         term_bar_layout.addWidget(term_label)
#         term_bar_layout.addStretch()
#         log_card_layout.addWidget(term_bar)

#         self.log = QTextEdit()
#         self.log.setReadOnly(True)
#         self.log.setMinimumHeight(180)
#         self.log.setFont(QFont("Consolas", 10))
#         self.log.setStyleSheet(
#             f"""
#             QTextEdit {{
#                 background-color: transparent;
#                 color: {TEXT_SECONDARY};
#                 border: none;
#                 padding: 16px;
#             }}
#             """
#         )
#         log_card_layout.addWidget(self.log, stretch=1)

#         body_layout.addWidget(log_card, stretch=1)
#         root.addWidget(body, stretch=1)

#         self.worker = None
#         self._log_placeholder()

#     # ------------------------------------------------------------------
#     def _log_placeholder(self):
#         self.log.setHtml(
#             f"<span style='color:{TEXT_MUTED};'>Waiting for files… select or drop an Excel file "
#             f"and a JSON/XML data file above, then hit Run.</span>"
#         )

#     def _update_run_state(self):
#         self.run_btn.setEnabled(bool(self.excel_zone.path and self.data_zone.path))

#     def _set_status(self, text, kind="info"):
#         color = {"success": SUCCESS, "error": ERROR, "info": ACCENT}[kind]
#         self.status_label.setText(f"  {text}")
#         self.status_label.setStyleSheet(
#             f"color: {color}; background-color: rgba(78,222,163,0.08); border: 1px solid {color}; "
#             f"border-radius: 10px; padding: 10px;"
#         )
#         self.status_label.show()

#     def _run(self):
#         save_path = None
#         if self.save_as_checkbox.isChecked():
#             base, ext = os.path.splitext(self.excel_zone.path)
#             suggested = f"{base}_filled{ext or '.xlsx'}"
#             chosen, _ = QFileDialog.getSaveFileName(
#                 self, "Save As — Choose destination", suggested,
#                 "Excel Files (*.xlsx *.xlsm)",
#             )
#             if not chosen:
#                 return  # user cancelled the Save As dialog - don't run
#             if not chosen.lower().endswith((".xlsx", ".xlsm")):
#                 chosen += ".xlsx"
#             save_path = chosen

#         self.run_btn.setEnabled(False)
#         self.run_btn.setText("Processing…")
#         self.status_label.hide()
#         self.log.clear()
#         self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Excel:</b> {self.excel_zone.path}")
#         self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Data:</b> {self.data_zone.path}")
#         if save_path:
#             self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Save as:</b> {save_path}")
#         self._append_log(f"<span style='color:{TEXT_MUTED}'>Reading sheets, detecting headers, matching data…</span><br>")

#         skip_unrelated = not self.fill_all_checkbox.isChecked()
#         self.worker = WorkerThread(
#             self.excel_zone.path, self.data_zone.path, skip_unrelated, save_path=save_path
#         )
#         self.worker.finished_ok.connect(self._on_success)
#         self.worker.finished_err.connect(self._on_error)
#         self.worker.start()

#     def _append_log(self, html):
#         self.log.append(html)

#     def _on_success(self, summary):
#         self.run_btn.setEnabled(True)
#         self.run_btn.setText("▶  Run — Fill Excel File")

#         self._append_log(
#             f"<span style='color:{TEXT_PRIMARY}'>Records found in data file: "
#             f"<b>{summary['records_processed']}</b></span><br>"
#         )
#         for sheet in summary["sheets"]:
#             self._append_log(f"<b style='color:{SECONDARY}'>■ {sheet['sheet']}</b>")
#             if sheet["skipped"]:
#                 self._append_log(
#                     f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;No matching headers — sheet left untouched.</span><br>"
#                 )
#                 continue
#             self._append_log(
#                 f"<span style='color:{TERTIARY}'>&nbsp;&nbsp;Rows filled: {sheet['rows_filled']} "
#                 f"(new: {sheet['rows_added']}, replaced duplicates: {sheet['rows_replaced']})</span>"
#             )
#             if sheet["matched_headers"]:
#                 self._append_log(f"<span style='color:{TEXT_SECONDARY}'>&nbsp;&nbsp;Matched columns:</span>")
#                 for header, key in sheet["matched_headers"].items():
#                     self._append_log(
#                         f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;&nbsp;&nbsp;'{header}' ← '{key}'</span>"
#                     )
#             if sheet["unmatched_headers"]:
#                 self._append_log(
#                     f"<span style='color:{ERROR}'>&nbsp;&nbsp;Filled with N/A (no matching data): "
#                     f"{', '.join(sheet['unmatched_headers'])}</span>"
#                 )
#             self._append_log("")

#         if summary.get("save_fallback"):
#             self._append_log(
#                 f"<b style='color:{ERROR}'>⚠ '{self.excel_zone.path}' was locked (probably open in Excel) — "
#                 f"saved a copy instead:</b>"
#             )
#             self._append_log(f"<b style='color:{TERTIARY}'>✓ Saved to: {summary['saved_to']}</b>")
#             self._set_status(
#                 "⚠ Original file was open/locked — saved a copy instead (see log for path).",
#                 kind="error",
#             )
#         else:
#             self._append_log(f"<b style='color:{TERTIARY}'>✓ Done — saved to: {summary['saved_to']}</b>")
#             self._set_status("✓ Excel file filled and saved successfully.", kind="success")

#     def _on_error(self, err_text):
#         self.run_btn.setEnabled(True)
#         self.run_btn.setText("▶  Run — Fill Excel File")
#         self._append_log(f"<span style='color:{ERROR}'><b>ERROR</b></span>")
#         self._append_log(f"<span style='color:{ERROR}'>{err_text}</span>")
#         self._set_status("✗ Something went wrong — see log for details.", kind="error")


# def main():
#     app = QApplication(sys.argv)
#     app.setStyle("Fusion")
#     if os.path.exists(LOGO_PATH):
#         app.setWindowIcon(QIcon(LOGO_PATH))
#     win = MainWindow()
#     win.show()
#     sys.exit(app.exec())


# if __name__ == "__main__":
#     main()





































"""
Excel Auto-Fill Bot - Desktop App

Pick (or drag & drop) an Excel file (multi-sheet, dynamic headers - nothing
hardcoded) and a JSON or XML data file (e.g. an Aadhaar-style API response).
The app fuzzy-matches the data's keys against each sheet's headers, fills
matching values into a brand new row of every sheet, writes "N/A" for
anything with no matching data, and overwrites the same Excel file.

- Har record (chahe duplicate hi kyu na ho) apni alag row leta hai - kuch bhi
  skip/merge nahi hota.
- By default koi bhi sheet skip nahi hoti - jis sheet mein match nahi milta
  wahan bhi ek row jaake N/A se bhar jaati hai. "Sirf unrelated sheets skip
  karo" checkbox se old (strict) behaviour wapas mil jata hai.
"""

import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent, QColor, QIcon, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from filler_core import fill_workbook

# ---------------------------------------------------------------------------
# Palette (matches the Excel Auto-Fill Bot design spec)
# ---------------------------------------------------------------------------
BG = "#0b1326"
SURFACE = "#0b1326"
PANEL = "#171f33"
PANEL_LIGHT = "#1e2338"
SURFACE_HIGH = "#222a3d"
SURFACE_HIGHEST = "#2d3449"
BORDER = "#464554"
BORDER_ACTIVE = "#8083ff"
ACCENT = "#c0c1ff"                 # primary
ACCENT_CONTAINER = "#8083ff"       # primary-container
ACCENT_HOVER = "#6f72e8"
ACCENT_SOFT = "#2a2d55"
SECONDARY = "#89ceff"
SECONDARY_CONTAINER = "#00a2e6"
TERTIARY = "#4edea3"
TERTIARY_CONTAINER = "#00885d"
TEXT_PRIMARY = "#dae2fd"
TEXT_SECONDARY = "#c7c4d7"
TEXT_MUTED = "#908fa0"
SUCCESS = "#4edea3"
ERROR = "#ffb4ab"

BODY_MAX_CONTENT_WIDTH = 1240  # inner content stays readable/centered on wide or maximized windows
BODY_MIN_MARGIN = 28  # smallest side margin, used once the window is narrower than the cap

ASSET_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(ASSET_DIR, "logo.png")
CHECK_ICON_PATH = os.path.join(ASSET_DIR, "check_icon.svg").replace("\\", "/")


def _checkbox_style() -> str:
    """
    Clean, modern checkbox: a rounded square that's a soft outline when
    unchecked and fills solid with a crisp checkmark when checked, with a
    subtle hover highlight. Qt's stylesheet engine takes over the indicator
    entirely once any indicator property is set, so the checkmark glyph is
    supplied explicitly via CHECK_ICON_PATH rather than left to the OS theme.
    """
    return f"""
        QCheckBox {{
            spacing: 12px;
            color: {TEXT_SECONDARY};
            font-size: 10.5pt;
            background: transparent;
            border: none;
            padding: 2px 0;
        }}
        QCheckBox::indicator {{
            width: 19px;
            height: 19px;
            border-radius: 6px;
            border: 1.5px solid {BORDER};
            background-color: {SURFACE_HIGH};
        }}
        QCheckBox::indicator:hover {{
            border-color: {BORDER_ACTIVE};
            background-color: {ACCENT_SOFT};
        }}
        QCheckBox::indicator:checked {{
            border-color: {ACCENT_CONTAINER};
            background-color: {ACCENT_CONTAINER};
            image: url({CHECK_ICON_PATH});
        }}
        QCheckBox::indicator:checked:hover {{
            border-color: {ACCENT_HOVER};
            background-color: {ACCENT_HOVER};
        }}
    """


class DropZone(QFrame):
    """Clickable + drag-and-drop file picker box (bento-grid card)."""

    def __init__(self, icon, icon_color, title, hint, extensions, badge_text, on_pick):
        super().__init__()
        self.extensions = extensions
        self.on_pick = on_pick
        self.path = None

        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(210)
        self.setMinimumWidth(280)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self._set_style(active=False)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(28)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(0, 0, 0, 110))
        self.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 28, 24, 28)
        layout.setSpacing(12)
        layout.setAlignment(Qt.AlignCenter)

        icon_wrap = QLabel(icon)
        icon_wrap.setAlignment(Qt.AlignCenter)
        icon_wrap.setFixedSize(64, 64)
        icon_wrap.setStyleSheet(
            f"font-size: 26px; background-color: {SURFACE_HIGH}; border: none; border-radius: 18px; color: {icon_color};"
        )

        icon_row = QHBoxLayout()
        icon_row.addStretch()
        icon_row.addWidget(icon_wrap)
        icon_row.addStretch()

        self.title_label = QLabel(title)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setWordWrap(True)
        self.title_label.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
        self.title_label.setMinimumHeight(24)
        self.title_label.setStyleSheet(f"""
QLabel {{
    color: {TEXT_PRIMARY};
    background: transparent;
    border: none;
}}
""")

        self.badge_label = QLabel(badge_text)
        self.badge_label.setAlignment(Qt.AlignCenter)
        self.badge_label.setWordWrap(True)

        self.hint_label = QLabel(hint)
        self.hint_label.setAlignment(Qt.AlignCenter)
        self.hint_label.setWordWrap(True)
        self.hint_label.setStyleSheet(f"""
QLabel {{
 color: {TEXT_MUTED};
 background: transparent;
 border: none;
 font-size:10pt;
}}
""")

        layout.addLayout(icon_row)
        layout.addSpacing(4)
        layout.addWidget(self.title_label)
        layout.addWidget(self.badge_label)
        layout.addWidget(self.hint_label)

        self._default_badge = badge_text
        self._default_hint = hint
        self._icon_color = icon_color
        self._style_badge(filled=False)

    def _style_badge(self, filled: bool):
        if filled:
            self.badge_label.setStyleSheet(f"""
QLabel {{
    color: {self._icon_color};
    background-color: rgba(255,255,255,0.06);
    border: none;
    border-radius: 999px;
    padding: 4px 14px;
    font-size: 10.5pt;
    font-weight: 600;
}}
""")
        else:
            self.badge_label.setStyleSheet(f"""
QLabel {{
    color: {TEXT_MUTED};
    background-color: rgba(255,255,255,0.03);
    border: 1px dashed {BORDER};
    border-radius: 999px;
    padding: 4px 14px;
    font-size: 10pt;
    font-style: italic;
}}
""")

    def _set_style(self, active: bool, filled: bool = False):
        if active:
            border, bg = BORDER_ACTIVE, ACCENT_SOFT
        elif filled:
            border, bg = TERTIARY, PANEL_LIGHT
        else:
            border, bg = BORDER, PANEL
        self.setStyleSheet(f"""
DropZone {{
    background-color: {bg};
    border: 2px solid {border};
    border-radius: 18px;
}}
QLabel {{
    border:none;
    background:transparent;
}}
""")

    def _valid_path(self, path: str) -> bool:
        return path.lower().endswith(tuple(self.extensions))

    def _apply_path(self, path: str):
        self.path = path
        self.badge_label.setText(os.path.basename(path))
        self._style_badge(filled=True)
        self.hint_label.setText("Click or drop again to replace")
        self._set_style(active=False, filled=True)
        self.on_pick()

    def reset(self):
        self.path = None
        self.badge_label.setText(self._default_badge)
        self._style_badge(filled=False)
        self.hint_label.setText(self._default_hint)
        self._set_style(active=False, filled=False)

    def mousePressEvent(self, event):
        filt = " ".join(f"*{e}" for e in self.extensions)
        path, _ = QFileDialog.getOpenFileName(self, "Select file", "", f"Supported Files ({filt})")
        if path and self._valid_path(path):
            self._apply_path(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            if self._valid_path(url.toLocalFile()):
                self._set_style(active=True)
                event.acceptProposedAction()
                return
        event.ignore()

    def dragLeaveEvent(self, event):
        self._set_style(active=False, filled=bool(self.path))

    def dropEvent(self, event: QDropEvent):
        url = event.mimeData().urls()[0]
        path = url.toLocalFile()
        if self._valid_path(path):
            self._apply_path(path)
            event.acceptProposedAction()
        else:
            self._set_style(active=False, filled=bool(self.path))
            event.ignore()


class WorkerThread(QThread):
    finished_ok = Signal(dict)
    finished_err = Signal(str)

    def __init__(self, excel_path, data_path, skip_unrelated_sheets, save_path=None):
        super().__init__()
        self.excel_path = excel_path
        self.data_path = data_path
        self.skip_unrelated_sheets = skip_unrelated_sheets
        self.save_path = save_path

    def run(self):
        try:
            summary = fill_workbook(
                self.excel_path,
                self.data_path,
                save_path=self.save_path,
                skip_unrelated_sheets=self.skip_unrelated_sheets,
            )
            self.finished_ok.emit(summary)
        except PermissionError as e:
            self.finished_err.emit(str(e))
        except Exception:
            self.finished_err.emit(traceback.format_exc())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Auto-Fill Bot")
        self.resize(940, 800)
        self.setMinimumSize(760, 640)
        self.setStyleSheet(f"QMainWindow {{ background-color: {BG}; }}")
        if os.path.exists(LOGO_PATH):
            self.setWindowIcon(QIcon(LOGO_PATH))

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---------- Header (AppBar) ----------
        header = QWidget()
        header.setStyleSheet(f"background-color: {SURFACE}; border-bottom: 1px solid {BORDER};")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(32, 26, 32, 22)
        header_layout.setSpacing(4)

        header_row = QHBoxLayout()
        badge = QLabel()
        badge.setFixedSize(48, 48)
        if os.path.exists(LOGO_PATH):
            badge.setPixmap(
                QPixmap(LOGO_PATH).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            )
        else:
            badge.setText("⚡")
            badge.setStyleSheet(
                f"background-color: {ACCENT_SOFT}; border-radius: 12px; font-size: 20px; color: {ACCENT};"
            )
        badge.setAlignment(Qt.AlignCenter)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        heading = QLabel("Excel Auto-Fill Bot")
        heading.setFont(QFont("Segoe UI", 21, QFont.Bold))
        heading.setStyleSheet(f"color: {TEXT_PRIMARY}; background: transparent; border: none;")
        title_box.addWidget(heading)

        subtitle = QLabel("Automated spreadsheet filling from JSON / XML")
        subtitle.setFont(QFont("Segoe UI", 9))
        subtitle.setStyleSheet(f"color: {TEXT_MUTED}; background: transparent; border: none;")
        title_box.addWidget(subtitle)

        header_row.addWidget(badge)
        header_row.addSpacing(10)
        header_row.addLayout(title_box)
        header_row.addStretch()

        self.refresh_btn = QPushButton("⟳  Refresh")
        self.refresh_btn.setCursor(Qt.PointingHandCursor)
        self.refresh_btn.setMinimumHeight(38)
        self.refresh_btn.setFont(QFont("Segoe UI", 10, QFont.DemiBold))
        self.refresh_btn.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {SURFACE_HIGH};
                color: {TEXT_SECONDARY};
                border: 1px solid {BORDER};
                border-radius: 10px;
                padding: 6px 16px;
            }}
            QPushButton:hover {{
                border-color: {BORDER_ACTIVE};
                color: {TEXT_PRIMARY};
                background-color: {ACCENT_SOFT};
            }}
            QPushButton:pressed {{
                background-color: {SURFACE_HIGHEST};
            }}
            """
        )
        self.refresh_btn.setToolTip("Clear the selected files, log and status to start over")
        self.refresh_btn.clicked.connect(self._refresh)
        header_row.addWidget(self.refresh_btn)

        header_layout.addLayout(header_row)
        root.addWidget(header)

        # ---------- Body ----------
        # `body` is the single widget under `root`, so it always fills the full
        # remaining width AND height automatically (no nested-layout stretch
        # ambiguity). To keep the content from stretching edge-to-edge on a
        # maximized / ultra-wide window, the left/right margins are recomputed
        # on every resize in `_update_body_margins` so the inner content stays
        # capped at BODY_MAX_CONTENT_WIDTH and centered, on any display size.
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setSpacing(20)
        self._body_layout = body_layout
        self._update_body_margins(self.width())

        # Drop zones (bento grid)
        picker_row = QHBoxLayout()
        picker_row.setSpacing(16)
        self.excel_zone = DropZone(
            "📊", TERTIARY, "Excel File",
            "Drag & drop .xlsx here\nor click to browse",
            [".xlsx", ".xlsm"], "No file selected", self._update_run_state,
        )
        self.data_zone = DropZone(
            "🗂", SECONDARY, "Data File (JSON / XML)",
            "Drag & drop .json / .xml here\nor click to browse",
            [".json", ".xml"], "No file selected", self._update_run_state,
        )
        picker_row.addWidget(self.excel_zone)
        picker_row.addWidget(self.data_zone)
        body_layout.addLayout(picker_row)

        # ---------- Options card ----------
        options_card = QFrame()
        options_card.setStyleSheet(
            f"QFrame {{ background-color: {PANEL}; border: 1px solid {BORDER}; border-radius: 14px; }}"
        )
        options_layout = QVBoxLayout(options_card)
        options_layout.setContentsMargins(20, 16, 20, 16)
        options_layout.setSpacing(12)

        options_title = QLabel("OPTIONS")
        options_title.setStyleSheet(
            f"color: {TEXT_MUTED}; font-size: 8.5pt; letter-spacing: 1px; "
            f"background: transparent; border: none;"
        )
        options_layout.addWidget(options_title)

        def _add_option(checkbox: QCheckBox, description: str):
            checkbox.setCursor(Qt.PointingHandCursor)
            checkbox.setStyleSheet(_checkbox_style())
            row = QVBoxLayout()
            row.setSpacing(2)
            row.setContentsMargins(0, 0, 0, 0)
            row.addWidget(checkbox)
            desc = QLabel(description)
            desc.setWordWrap(True)
            desc.setStyleSheet(
                f"color: {TEXT_MUTED}; font-size: 9pt; background: transparent; "
                f"border: none; margin-left: 31px;"
            )
            row.addWidget(desc)
            options_layout.addLayout(row)

        self.fill_all_checkbox = QCheckBox("Fill every sheet, even unrelated ones")
        self.fill_all_checkbox.setChecked(False)
        _add_option(
            self.fill_all_checkbox,
            "Sheets with no matching columns still get a new row, filled entirely with N/A.",
        )

        divider = QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet(f"background-color: {BORDER}; border: none;")
        options_layout.addWidget(divider)

        self.save_as_checkbox = QCheckBox("Save as a new Excel file")
        self.save_as_checkbox.setChecked(False)
        _add_option(
            self.save_as_checkbox,
            "Keeps the original workbook untouched and writes the result to a file you choose.",
        )

        body_layout.addWidget(options_card)

        # Run button
        self.run_btn = QPushButton("▶  Run — Fill Excel File")
        self.run_btn.setEnabled(False)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.setMinimumHeight(56)
        self.run_btn.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
        self.run_btn.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {ACCENT_CONTAINER};
                color: #07006c;
                border-radius: 12px;
                border: none;
            }}
            QPushButton:disabled {{
                background-color: {PANEL_LIGHT};
                color: {TEXT_MUTED};
            }}
            QPushButton:hover:!disabled {{
                background-color: {ACCENT_HOVER};
                color: #ffffff;
            }}
            QPushButton:pressed:!disabled {{
                background-color: #4f5bd1;
                color: #ffffff;
            }}
            """
        )
        self.run_btn.clicked.connect(self._run)
        body_layout.addWidget(self.run_btn)

        # Status pill
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Segoe UI", 10, QFont.Medium))
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setMinimumHeight(40)
        self.status_label.hide()
        body_layout.addWidget(self.status_label)

        # Activity log (terminal style)
        log_card = QFrame()
        log_card.setStyleSheet(
            f"QFrame {{ background-color: {PANEL}; border: 1px solid {BORDER}; border-radius: 14px; }}"
        )
        log_card_layout = QVBoxLayout(log_card)
        log_card_layout.setContentsMargins(0, 0, 0, 0)
        log_card_layout.setSpacing(0)

        term_bar = QWidget()
        term_bar.setStyleSheet(
            f"background-color: {SURFACE_HIGH}; border-top-left-radius: 14px; border-top-right-radius: 14px;"
        )
        term_bar_layout = QHBoxLayout(term_bar)
        term_bar_layout.setContentsMargins(14, 8, 14, 8)
        for dot_color in (ERROR, SECONDARY, TERTIARY):
            dot = QLabel()
            dot.setFixedSize(10, 10)
            dot.setStyleSheet(f"background-color: {dot_color}; border-radius: 5px;")
            term_bar_layout.addWidget(dot)
            term_bar_layout.addSpacing(4)
        term_bar_layout.addSpacing(10)
        term_label = QLabel("TERMINAL — PROCESS_MAPPING")
        term_label.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 8.5pt; letter-spacing: 1px;")
        term_bar_layout.addWidget(term_label)
        term_bar_layout.addStretch()
        log_card_layout.addWidget(term_bar)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(180)
        self.log.setFont(QFont("Consolas", 10))
        self.log.setStyleSheet(
            f"""
            QTextEdit {{
                background-color: transparent;
                color: {TEXT_SECONDARY};
                border: none;
                padding: 16px;
            }}
            """
        )
        log_card_layout.addWidget(self.log, stretch=1)

        body_layout.addWidget(log_card, stretch=1)
        root.addWidget(body, stretch=1)

        self.worker = None
        self._log_placeholder()

    # ------------------------------------------------------------------
    def _update_body_margins(self, window_width: int):
        side = max(BODY_MIN_MARGIN, (window_width - BODY_MAX_CONTENT_WIDTH) // 2)
        self._body_layout.setContentsMargins(side, 28, side, 28)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_body_margins(event.size().width())

    def _log_placeholder(self):
        self.log.setHtml(
            f"<span style='color:{TEXT_MUTED};'>Waiting for files… select or drop an Excel file "
            f"and a JSON/XML data file above, then hit Run.</span>"
        )

    def _update_run_state(self):
        self.run_btn.setEnabled(bool(self.excel_zone.path and self.data_zone.path))

    def _refresh(self):
        """Reset the whole form back to its starting state for a fresh run."""
        if self.worker is not None and self.worker.isRunning():
            return  # don't reset out from under an in-progress run
        self.excel_zone.reset()
        self.data_zone.reset()
        self.fill_all_checkbox.setChecked(False)
        self.save_as_checkbox.setChecked(False)
        self.status_label.hide()
        self.log.clear()
        self._log_placeholder()
        self.run_btn.setEnabled(False)
        self.run_btn.setText("▶  Run — Fill Excel File")

    def _set_status(self, text, kind="info"):
        color = {"success": SUCCESS, "error": ERROR, "info": ACCENT}[kind]
        self.status_label.setText(f"  {text}")
        self.status_label.setStyleSheet(
            f"color: {color}; background-color: rgba(78,222,163,0.08); border: 1px solid {color}; "
            f"border-radius: 10px; padding: 10px;"
        )
        self.status_label.show()

    def _run(self):
        save_path = None
        if self.save_as_checkbox.isChecked():
            base, ext = os.path.splitext(self.excel_zone.path)
            suggested = f"{base}_filled{ext or '.xlsx'}"
            chosen, _ = QFileDialog.getSaveFileName(
                self, "Save As — Choose destination", suggested,
                "Excel Files (*.xlsx *.xlsm)",
            )
            if not chosen:
                return  # user cancelled the Save As dialog - don't run
            if not chosen.lower().endswith((".xlsx", ".xlsm")):
                chosen += ".xlsx"
            save_path = chosen

        self.run_btn.setEnabled(False)
        self.run_btn.setText("Processing…")
        self.status_label.hide()
        self.log.clear()
        self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Excel:</b> {self.excel_zone.path}")
        self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Data:</b> {self.data_zone.path}")
        if save_path:
            self._append_log(f"<b style='color:{TEXT_PRIMARY}'>Save as:</b> {save_path}")
        self._append_log(f"<span style='color:{TEXT_MUTED}'>Reading sheets, detecting headers, matching data…</span><br>")

        skip_unrelated = not self.fill_all_checkbox.isChecked()
        self.worker = WorkerThread(
            self.excel_zone.path, self.data_zone.path, skip_unrelated, save_path=save_path
        )
        self.worker.finished_ok.connect(self._on_success)
        self.worker.finished_err.connect(self._on_error)
        self.worker.start()

    def _append_log(self, html):
        self.log.append(html)

    def _on_success(self, summary):
        self.run_btn.setEnabled(True)
        self.run_btn.setText("▶  Run — Fill Excel File")

        self._append_log(
            f"<span style='color:{TEXT_PRIMARY}'>Records found in data file: "
            f"<b>{summary['records_processed']}</b></span><br>"
        )
        for sheet in summary["sheets"]:
            self._append_log(f"<b style='color:{SECONDARY}'>■ {sheet['sheet']}</b>")
            if sheet["skipped"]:
                self._append_log(
                    f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;No matching headers — sheet left untouched.</span><br>"
                )
                continue
            self._append_log(
                f"<span style='color:{TERTIARY}'>&nbsp;&nbsp;Rows filled: {sheet['rows_filled']} "
                f"(new: {sheet['rows_added']}, replaced duplicates: {sheet['rows_replaced']})</span>"
            )
            if sheet["matched_headers"]:
                self._append_log(f"<span style='color:{TEXT_SECONDARY}'>&nbsp;&nbsp;Matched columns:</span>")
                for header, key in sheet["matched_headers"].items():
                    self._append_log(
                        f"<span style='color:{TEXT_MUTED}'>&nbsp;&nbsp;&nbsp;&nbsp;'{header}' ← '{key}'</span>"
                    )
            if sheet["unmatched_headers"]:
                self._append_log(
                    f"<span style='color:{ERROR}'>&nbsp;&nbsp;Filled with N/A (no matching data): "
                    f"{', '.join(sheet['unmatched_headers'])}</span>"
                )
            self._append_log("")

        if summary.get("save_fallback"):
            self._append_log(
                f"<b style='color:{ERROR}'>⚠ '{self.excel_zone.path}' was locked (probably open in Excel) — "
                f"saved a copy instead:</b>"
            )
            self._append_log(f"<b style='color:{TERTIARY}'>✓ Saved to: {summary['saved_to']}</b>")
            self._set_status(
                "⚠ Original file was open/locked — saved a copy instead (see log for path).",
                kind="error",
            )
        else:
            self._append_log(f"<b style='color:{TERTIARY}'>✓ Done — saved to: {summary['saved_to']}</b>")
            self._set_status("✓ Excel file filled and saved successfully.", kind="success")

        self._append_log(
            f"<span style='color:{TEXT_MUTED}'>Tip: hit ⟳ Refresh (top right) to clear the "
            f"selected files and run another record.</span>"
        )

    def _on_error(self, err_text):
        self.run_btn.setEnabled(True)
        self.run_btn.setText("▶  Run — Fill Excel File")
        self._append_log(f"<span style='color:{ERROR}'><b>ERROR</b></span>")
        self._append_log(f"<span style='color:{ERROR}'>{err_text}</span>")
        self._set_status("✗ Something went wrong — see log for details.", kind="error")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    if os.path.exists(LOGO_PATH):
        app.setWindowIcon(QIcon(LOGO_PATH))
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
