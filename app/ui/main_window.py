from __future__ import annotations

import sys
import json
from pathlib import Path
from datetime import datetime
from typing import List

import pandas as pd
from PyQt5.QtCore import QThread
from PyQt5.QtGui import QColor, QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
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
from app.io.excel_writer import RESULT_COLUMNS
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
        self.forbidden_rules_list: QListWidget | None = None
        self.forbidden_term_input: QLineEdit | None = None
        self.forbidden_category_input: QLineEdit | None = None
        self.forbidden_severity_combo: QComboBox | None = None
        self.tab_forbidden_result: QTextEdit | None = None

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
        self.btn_one_click = QPushButton("원클릭 전체 진단")
        self.btn_one_click.clicked.connect(self._start_one_click_analysis)
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
        left_layout.addWidget(self.btn_one_click)
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
        self.tab_forbidden = self._build_forbidden_tab()
        self.tab_suggestion = QTextEdit()
        self.tab_settings = QTextEdit()
        self.tab_logs = QTextEdit()
        self.btn_save_settings = QPushButton("설정 저장")
        self.btn_save_settings.clicked.connect(self._save_settings_from_tab)
        self.tab_settings.setPlainText(json.dumps(self.settings, ensure_ascii=False, indent=2))
        self.tabs.addTab(self.tab_byte, "Byte 분석")
        self.tabs.addTab(self.tab_forbidden, "금지표현 관리/검출")
        self.tabs.addTab(self.tab_suggestion, "대체표현 추천")
        self.tabs.addTab(self.tab_settings, "설정")
        self.tabs.addTab(self.tab_logs, "로그")
        self.tab_byte.setReadOnly(True)
        self.tab_suggestion.setReadOnly(True)
        self.tab_logs.setReadOnly(True)
        self.tab_settings.setReadOnly(False)
        root_layout.addWidget(self.tabs)
        root_layout.addWidget(self.btn_save_settings)

    def _build_forbidden_tab(self) -> QWidget:
        wrapper = QWidget()
        layout = QVBoxLayout(wrapper)

        splitter = QSplitter()
        layout.addWidget(splitter)

        manager_panel = QWidget()
        manager_layout = QVBoxLayout(manager_panel)
        manager_layout.addWidget(QLabel("금지/주의 표현 규칙 관리"))

        self.forbidden_rules_list = QListWidget()
        self.forbidden_rules_list.setSelectionMode(QListWidget.ExtendedSelection)
        manager_layout.addWidget(self.forbidden_rules_list)

        self.forbidden_term_input = QLineEdit()
        self.forbidden_term_input.setPlaceholderText("표현(term)")
        self.forbidden_category_input = QLineEdit()
        self.forbidden_category_input.setPlaceholderText("카테고리(category)")
        self.forbidden_severity_combo = QComboBox()
        self.forbidden_severity_combo.addItems(["forbidden", "review"])

        manager_layout.addWidget(QLabel("표현"))
        manager_layout.addWidget(self.forbidden_term_input)
        manager_layout.addWidget(QLabel("카테고리"))
        manager_layout.addWidget(self.forbidden_category_input)
        manager_layout.addWidget(QLabel("심각도"))
        manager_layout.addWidget(self.forbidden_severity_combo)

        button_row = QHBoxLayout()
        btn_add = QPushButton("규칙 추가")
        btn_add.clicked.connect(self._add_forbidden_rule)
        btn_remove = QPushButton("선택 삭제")
        btn_remove.clicked.connect(self._remove_selected_forbidden_rules)
        btn_reload = QPushButton("다시 불러오기")
        btn_reload.clicked.connect(self._reload_forbidden_rules)
        btn_save = QPushButton("금지어 저장")
        btn_save.clicked.connect(self._save_forbidden_rules)
        button_row.addWidget(btn_add)
        button_row.addWidget(btn_remove)
        button_row.addWidget(btn_reload)
        button_row.addWidget(btn_save)
        manager_layout.addLayout(button_row)

        result_panel = QWidget()
        result_layout = QVBoxLayout(result_panel)
        result_layout.addWidget(QLabel("최근 분석 결과(금지표현)"))
        self.tab_forbidden_result = QTextEdit()
        self.tab_forbidden_result.setReadOnly(True)
        result_layout.addWidget(self.tab_forbidden_result)

        splitter.addWidget(manager_panel)
        splitter.addWidget(result_panel)
        splitter.setSizes([700, 500])

        self._refresh_forbidden_rules_list()
        return wrapper

    def _analyze_dataframe_rows(self, df: pd.DataFrame, target_cols: List[str], item_type: str) -> pd.DataFrame:
        byte_counter = ByteCounter(ByteCounterConfig(newline_bytes=int(self.settings["newline_bytes"])))
        rule_engine = RuleEngine(self.forbidden, self.allowlist, self.categories)
        suggestion_engine = SuggestionEngine(self.suggestions)
        analyzer = Analyzer(byte_counter, rule_engine, suggestion_engine)
        byte_limit = int(self.settings["byte_limits"].get(item_type, 500))

        records = []
        for _, row in df.iterrows():
            row_hits = []
            row_suggestions = []
            row_notes = []
            max_bytes = 0
            overflow = "N"
            verdict = "허용"
            for col in target_cols:
                text = row.get(col, "")
                analyzed = analyzer.analyze_cell(col, "" if pd.isna(text) else str(text), byte_limit)
                max_bytes = max(max_bytes, analyzed.byte_count)
                overflow = "Y" if analyzed.overflow_yn == "Y" else overflow
                row_hits.extend(analyzed.hit_terms)
                if analyzed.suggested_rewrite:
                    row_suggestions.append(analyzed.suggested_rewrite)
                if analyzed.review_note:
                    row_notes.append(analyzed.review_note)
                if analyzed.verdict == "확정 위반":
                    verdict = "확정 위반"
                elif analyzed.verdict == "검토 필요" and verdict != "확정 위반":
                    verdict = "검토 필요"

            unique_hits = []
            for hit in row_hits:
                if hit not in unique_hits:
                    unique_hits.append(hit)
            unique_suggestions = []
            for suggestion in row_suggestions:
                if suggestion not in unique_suggestions:
                    unique_suggestions.append(suggestion)
            unique_notes = []
            for note in row_notes:
                if note not in unique_notes:
                    unique_notes.append(note)

            records.append(
                {
                    "byte_count": max_bytes,
                    "byte_limit": byte_limit,
                    "overflow_yn": overflow,
                    "verdict": verdict,
                    "hit_terms": ", ".join(unique_hits),
                    "suggested_rewrite": "; ".join(unique_suggestions),
                    "review_note": "; ".join(unique_notes),
                }
            )

        return pd.DataFrame(records)

    def _start_one_click_analysis(self) -> None:
        if self.current_file is None:
            QMessageBox.warning(self, "경고", "먼저 Excel 파일을 선택하세요.")
            return

        try:
            sheets = self.reader.list_sheets(self.current_file)
            if not sheets:
                QMessageBox.warning(self, "경고", "분석할 시트가 없습니다.")
                return

            merged_frames = []
            total_sheets = len(sheets)
            for idx, sheet in enumerate(sheets, start=1):
                sheet_df = self.reader.read_full(self.current_file, sheet, self.header_spin.value())
                if sheet_df.empty:
                    continue
                target_cols = [str(col) for col in sheet_df.columns]
                result_df = self._analyze_dataframe_rows(
                    sheet_df,
                    target_cols,
                    self.item_type_combo.currentText(),
                )
                combined = sheet_df.copy()
                for col in RESULT_COLUMNS:
                    combined[col] = result_df[col]
                combined["sheet_name"] = sheet
                merged_frames.append(combined)
                self.progress.setValue(int((idx / total_sheets) * 100))

            if not merged_frames:
                QMessageBox.warning(self, "경고", "분석 가능한 데이터가 없습니다.")
                return

            final_df = pd.concat(merged_frames, ignore_index=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = self.current_file.with_name(f"{self.current_file.stem}_원클릭진단_{stamp}.xlsx")

            summary = final_df["verdict"].value_counts(dropna=False).rename_axis("verdict").reset_index(name="count")
            violations = final_df[final_df["verdict"] == "확정 위반"].copy()
            reviews = final_df[final_df["verdict"] == "검토 필요"].copy()
            settings_df = pd.DataFrame([{"key": k, "value": str(v)} for k, v in self.settings.items()])

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                final_df.to_excel(writer, sheet_name="분석결과", index=False)
                summary.to_excel(writer, sheet_name="요약", index=False)
                violations.to_excel(writer, sheet_name="위반목록", index=False)
                reviews.to_excel(writer, sheet_name="검토필요목록", index=False)
                settings_df.to_excel(writer, sheet_name="설정 스냅샷", index=False)

            csv_path = output_path.with_suffix(".csv")
            final_df.to_csv(csv_path, index=False, encoding="utf-8-sig")

            self.tab_logs.setPlainText(f"원클릭 전체 진단 완료\n엑셀: {output_path}\nCSV: {csv_path}")
            if self.tab_forbidden_result is not None:
                self.tab_forbidden_result.setPlainText(
                    final_df[["verdict", "hit_terms", "review_note"]].head(100).to_string(index=False)
                )
            QMessageBox.information(self, "완료", f"원클릭 전체 진단 완료:\n{output_path}")
        except Exception as exc:  # pylint: disable=broad-except
            QMessageBox.critical(self, "오류", f"원클릭 전체 진단 실패: {exc}")

    def _refresh_forbidden_rules_list(self) -> None:
        if self.forbidden_rules_list is None:
            return
        self.forbidden_rules_list.clear()
        rules = self.forbidden.get("rules", [])
        for rule in rules:
            term = str(rule.get("term", ""))
            category = str(rule.get("category", ""))
            severity = str(rule.get("severity", "review"))
            item = QListWidgetItem(f"{term} | {category} | {severity}")
            item.setData(32, {"term": term, "category": category, "severity": severity})
            self.forbidden_rules_list.addItem(item)

    def _add_forbidden_rule(self) -> None:
        if not self.forbidden_term_input or not self.forbidden_category_input or not self.forbidden_severity_combo:
            return
        term = self.forbidden_term_input.text().strip()
        category = self.forbidden_category_input.text().strip()
        severity = self.forbidden_severity_combo.currentText()
        if not term or not category:
            QMessageBox.warning(self, "경고", "표현과 카테고리를 모두 입력하세요.")
            return

        rules = self.forbidden.setdefault("rules", [])
        duplicate = next((r for r in rules if r.get("term") == term and r.get("category") == category), None)
        if duplicate is not None:
            QMessageBox.warning(self, "경고", "동일한 표현/카테고리 규칙이 이미 존재합니다.")
            return

        rules.append({"term": term, "category": category, "severity": severity})
        self._refresh_forbidden_rules_list()
        self.forbidden_term_input.clear()
        self.forbidden_category_input.clear()

    def _remove_selected_forbidden_rules(self) -> None:
        if self.forbidden_rules_list is None:
            return
        selected = self.forbidden_rules_list.selectedItems()
        if not selected:
            QMessageBox.warning(self, "경고", "삭제할 규칙을 선택하세요.")
            return
        selected_keys = {(
            item.data(32).get("term", ""),
            item.data(32).get("category", ""),
            item.data(32).get("severity", "review"),
        ) for item in selected}
        self.forbidden["rules"] = [
            rule
            for rule in self.forbidden.get("rules", [])
            if (rule.get("term", ""), rule.get("category", ""), rule.get("severity", "review")) not in selected_keys
        ]
        self._refresh_forbidden_rules_list()

    def _reload_forbidden_rules(self) -> None:
        self.forbidden = self.config_manager.load_json("forbidden_rules.json")
        self._refresh_forbidden_rules_list()
        QMessageBox.information(self, "완료", "저장된 금지어 규칙을 다시 불러왔습니다.")

    def _save_forbidden_rules(self) -> None:
        self.config_manager.save_json("forbidden_rules.json", self.forbidden)
        QMessageBox.information(self, "저장", "금지어 규칙 저장 완료")

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
        if self.tab_forbidden_result is not None:
            self.tab_forbidden_result.setPlainText(
                result_df[["verdict", "hit_terms", "review_note"]].head(100).to_string(index=False)
            )
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
