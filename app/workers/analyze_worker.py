from __future__ import annotations

from typing import Dict, List

import pandas as pd
from PyQt5.QtCore import QObject, pyqtSignal

from app.core.analyzer import Analyzer


class AnalyzeWorker(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal(pd.DataFrame)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(
        self,
        df: pd.DataFrame,
        target_columns: List[str],
        item_type: str,
        byte_limits: Dict[str, int],
        analyzer: Analyzer,
        chunk_size: int = 300,
    ) -> None:
        super().__init__()
        self.df = df
        self.target_columns = target_columns
        self.item_type = item_type
        self.byte_limits = byte_limits
        self.analyzer = analyzer
        self.chunk_size = chunk_size
        self._cancel_requested = False

    def cancel(self) -> None:
        self._cancel_requested = True

    def run(self) -> None:
        try:
            records = []
            total = len(self.df.index)
            current = 0
            byte_limit = int(self.byte_limits.get(self.item_type, 500))
            for _, row in self.df.iterrows():
                if self._cancel_requested:
                    self.cancelled.emit()
                    return
                row_hits = []
                row_suggestions = []
                row_notes = []
                max_bytes = 0
                overflow = "N"
                verdict = "허용"
                for col in self.target_columns:
                    text = row.get(col, "")
                    analyzed = self.analyzer.analyze_cell(col, "" if pd.isna(text) else str(text), byte_limit)
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
                for h in row_hits:
                    if h not in unique_hits:
                        unique_hits.append(h)
                unique_suggestions = []
                for s in row_suggestions:
                    if s not in unique_suggestions:
                        unique_suggestions.append(s)
                unique_notes = []
                for n in row_notes:
                    if n not in unique_notes:
                        unique_notes.append(n)

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
                current += 1
                self.progress.emit(int((current / total) * 100))
            self.finished.emit(pd.DataFrame(records))
        except Exception as exc:  # pylint: disable=broad-except
            self.failed.emit(str(exc))
