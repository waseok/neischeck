from __future__ import annotations

import sys
import json
from pathlib import Path
from typing import List

import pandas as pd
from PyQt5.QtCore import QThread
from PyQt5.QtGui import QColor, QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.config_manager import ConfigManager
from app.core.analyzer import Analyzer
from app.core.byte_counter import ByteCounter, ByteCounterConfig
from app.core.rule_engine import RuleEngine
from app.core.suggestion_engine import SuggestionEngine
from app.io.excel_reader import ExcelReader
from app.io.excel_writer import ExcelWriter
from app.utils.logging_setup import setup_logger
from app.workers.analyze_worker import AnalyzeWorker


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("생기부 특기사항 자동 검토기")
        self.resize(1500, 900)

        self.config_manager = ConfigManager()
        self.settings = self.config_manager.load_json("settings.json")
        self.forbidden = self.config_manager.load_json("forbidden_rules.json")
        self.allowlist = self.config_manager.load_json("allowlist.json")
        self.categories = self.config_manager.load_json("category_rules.json")
        self.suggestions = self.config_manager.load_json("suggestion_rules.json")
        self.logger = setup_logger(Path.cwd() / "data" / "logs", self.settings.get("log_raw_text", False))

        self.reader = ExcelReader()
        self.writer = ExcelWriter()
        self.current_file: Path | None = None
        self.current_df = pd.DataFrame()
        self.result_df = pd.DataFrame()
        self.thread: QThread | None = None
        self.worker: AnalyzeWorker | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)

        top_layout = QHBoxLayout()
        root_layout.addLayout(top_layout)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        self.btn_file = QPushButton("Excel 파일 선택")
        self.btn_file.clicked.connect(self._select_file)
        self.sheet_combo = QComboBox()
        self.header_spin = QSpinBox()
        self.header_spin.setMinimum(1)
        self.header_spin.setValue(1)
        self.header_spin.valueChanged.connect(self._refresh_preview)
        self.target_columns = QListWidget()
        self.id_columns = QListWidget()
        self.item_type_combo = QComboBox()
        self.item_type_combo.addItems(list(self.settings.get("byte_limits", {}).keys()))
        self.btn_analyze = QPushButton("분석 시작")
        self.btn_analyze.clicked.connect(self._start_analysis)
        self.btn_cancel = QPushButton("분석 취소")
        self.btn_cancel.clicked.connect(self._cancel_analysis)
        self.progress = QProgressBar()
        self.progress.setValue(0)

        left_layout.addWidget(self.btn_file)
        left_layout.addWidget(QLabel("시트 선택"))
        left_layout.addWidget(self.sheet_combo)
        left_layout.addWidget(QLabel("헤더 행"))
        left_layout.addWidget(self.header_spin)
        left_layout.addWidget(QLabel("검토 대상 열(복수 선택)"))
        left_layout.addWidget(self.target_columns)
        left_layout.addWidget(QLabel("학생 식별용 열(복수 선택)"))
        left_layout.addWidget(self.id_columns)
        left_layout.addWidget(QLabel("항목 유형(Byte 제한)"))
        left_layout.addWidget(self.item_type_combo)
        left_layout.addWidget(self.btn_analyze)
        left_layout.addWidget(self.btn_cancel)
        left_layout.addWidget(self.progress)

        self.table = QTableView()
        self.table.clicked.connect(self._on_table_clicked)
        self.detail = QTextEdit()
        self.detail.setReadOnly(True)

        top_layout.addWidget(left_panel, 2)
        top_layout.addWidget(self.table, 5)
        top_layout.addWidget(self.detail, 3)

        self.tabs = QTabWidget()
        self.tab_byte = QTextEdit()
        self.tab_forbidden = QTextEdit()
        self.tab_suggestion = QTextEdit()
        self.tab_settings = QTextEdit()
        self.tab_logs = QTextEdit()
        self.btn_save_settings = QPushButton("설정 저장")
        self.btn_save_settings.clicked.connect(self._save_settings_from_tab)
        self.tab_settings.setPlainText(json.dumps(self.settings, ensure_ascii=False, indent=2))
        for name, widget in [
            ("Byte 분석", self.tab_byte),
            ("금지표현 검출", self.tab_forbidden),
            ("대체표현 추천", self.tab_suggestion),
            ("설정", self.tab_settings),
            ("로그", self.tab_logs),
        ]:
            widget.setReadOnly(True)
            self.tabs.addTab(widget, name)
        self.tab_settings.setReadOnly(False)
        root_layout.addWidget(self.tabs)
        root_layout.addWidget(self.btn_save_settings)

    def _select_file(self) -> None:
        file_name, _ = QFileDialog.getOpenFileName(self, "Excel 선택", "", "Excel Files (*.xlsx *.xlsm)")
        if not file_name:
            return
        self.current_file = Path(file_name)
        sheets = self.reader.list_sheets(self.current_file)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheets)
        self.sheet_combo.currentIndexChanged.connect(self._refresh_preview)
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        if not self.current_file or not self.sheet_combo.currentText():
            return
        self.current_df = self.reader.read_preview(
            self.current_file,
            self.sheet_combo.currentText(),
            self.header_spin.value(),
            int(self.settings["ui"]["preview_rows"]),
        )
        self._set_table_dataframe(self.current_df)
        self.target_columns.clear()
        self.id_columns.clear()
        for col in self.current_df.columns:
            item = QListWidgetItem(str(col))
            item.setCheckState(0)
            self.target_columns.addItem(item)
            id_item = QListWidgetItem(str(col))
            id_item.setCheckState(0)
            self.id_columns.addItem(id_item)

    def _set_table_dataframe(self, df: pd.DataFrame, result_df: pd.DataFrame | None = None) -> None:
        model = QStandardItemModel()
        model.setColumnCount(len(df.columns))
        model.setHorizontalHeaderLabels([str(c) for c in df.columns])
        for r_idx, row in df.iterrows():
            items: List[QStandardItem] = []
            for c_idx, col in enumerate(df.columns):
                item = QStandardItem("" if pd.isna(row[col]) else str(row[col]))
                if result_df is not None and r_idx < len(result_df.index):
                    verdict = result_df.iloc[r_idx]["verdict"]
                    if verdict == "확정 위반":
                        item.setBackground(QColor(255, 128, 128))
                    elif verdict == "검토 필요":
                        item.setBackground(QColor(255, 235, 130))
                    else:
                        item.setBackground(QColor(200, 255, 200))
                items.append(item)
            model.appendRow(items)
        self.table.setModel(model)

    def _selected_target_columns(self) -> List[str]:
        selected = []
        for i in range(self.target_columns.count()):
            item = self.target_columns.item(i)
            if item.checkState() == 2:
                selected.append(item.text())
        return selected

    def _start_analysis(self) -> None:
        if self.current_file is None:
            QMessageBox.warning(self, "경고", "먼저 Excel 파일을 선택하세요.")
            return
        target_cols = self._selected_target_columns()
        if not target_cols:
            QMessageBox.warning(self, "경고", "검토 대상 열을 선택하세요.")
            return

        full_df = self.reader.read_full(self.current_file, self.sheet_combo.currentText(), self.header_spin.value())
        byte_counter = ByteCounter(ByteCounterConfig(newline_bytes=int(self.settings["newline_bytes"])))
        rule_engine = RuleEngine(self.forbidden, self.allowlist, self.categories)
        suggestion_engine = SuggestionEngine(self.suggestions)
        analyzer = Analyzer(byte_counter, rule_engine, suggestion_engine)

        self.thread = QThread()
        self.worker = AnalyzeWorker(
            full_df,
            target_cols,
            self.item_type_combo.currentText(),
            self.settings["byte_limits"],
            analyzer,
            int(self.settings["ui"]["row_chunk_size"]),
        )
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self._on_analysis_finished)
        self.worker.failed.connect(self._on_analysis_failed)
        self.worker.cancelled.connect(self._on_analysis_cancelled)
        self.thread.start()

    def _cancel_analysis(self) -> None:
        if self.worker is not None:
            self.worker.cancel()

    def _on_analysis_finished(self, result_df: pd.DataFrame) -> None:
        self.result_df = result_df
        self.tab_byte.setPlainText(result_df[["byte_count", "byte_limit", "overflow_yn"]].head(100).to_string(index=False))
        self.tab_forbidden.setPlainText(result_df[["verdict", "hit_terms", "review_note"]].head(100).to_string(index=False))
        self.tab_suggestion.setPlainText(result_df[["hit_terms", "suggested_rewrite"]].head(100).to_string(index=False))
        self.tab_logs.setPlainText("분석 완료")

        if self.current_file is not None:
            output_path = self.writer.build_output_path(self.current_file)
            source_df = self.reader.read_full(self.current_file, self.sheet_combo.currentText(), self.header_spin.value())
            self.writer.save_results(source_df, result_df, output_path, self.settings)
            self.writer.save_csv(result_df, output_path)
            QMessageBox.information(self, "완료", f"결과 파일 저장 완료:\n{output_path}")
            self._set_table_dataframe(self.current_df, result_df.head(len(self.current_df.index)))

        if self.thread:
            self.thread.quit()
            self.thread.wait()

    def _on_analysis_failed(self, message: str) -> None:
        self.tab_logs.setPlainText(f"오류: {message}")
        QMessageBox.critical(self, "오류", message)
        if self.thread:
            self.thread.quit()
            self.thread.wait()

    def _on_analysis_cancelled(self) -> None:
        self.tab_logs.setPlainText("분석 취소됨")
        if self.thread:
            self.thread.quit()
            self.thread.wait()

    def _on_table_clicked(self) -> None:
        index = self.table.currentIndex()
        if not index.isValid():
            return
        value = self.table.model().data(index)
        self.detail.setPlainText(f"선택 셀 값:\n{value}")

    def _save_settings_from_tab(self) -> None:
        try:
            payload = json.loads(self.tab_settings.toPlainText())
            self.settings = payload
            self.config_manager.save_json("settings.json", payload)
            QMessageBox.information(self, "저장", "설정 저장 완료")
        except Exception as exc:  # pylint: disable=broad-except
            QMessageBox.warning(self, "오류", f"설정 저장 실패: {exc}")


def run_app() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
