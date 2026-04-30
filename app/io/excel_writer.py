from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict

import pandas as pd


RESULT_COLUMNS = [
    "byte_count",
    "byte_limit",
    "overflow_yn",
    "verdict",
    "hit_terms",
    "suggested_rewrite",
    "review_note",
]


class ExcelWriter:
    def build_output_path(self, source_path: Path) -> Path:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return source_path.with_name(f"{source_path.stem}_검토결과_{stamp}.xlsx")

    def save_results(
        self,
        source_df: pd.DataFrame,
        result_df: pd.DataFrame,
        output_path: Path,
        settings_snapshot: Dict[str, object],
    ) -> None:
        merged = source_df.copy()
        for col in RESULT_COLUMNS:
            merged[col] = result_df[col]

        summary = result_df["verdict"].value_counts(dropna=False).rename_axis("verdict").reset_index(name="count")
        violations = result_df[result_df["verdict"] == "확정 위반"].copy()
        reviews = result_df[result_df["verdict"] == "검토 필요"].copy()
        settings_df = pd.DataFrame(
            [{"key": k, "value": str(v)} for k, v in settings_snapshot.items()]
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            merged.to_excel(writer, sheet_name="분석결과", index=False)
            summary.to_excel(writer, sheet_name="요약", index=False)
            violations.to_excel(writer, sheet_name="위반목록", index=False)
            reviews.to_excel(writer, sheet_name="검토필요목록", index=False)
            settings_df.to_excel(writer, sheet_name="설정 스냅샷", index=False)

    def save_csv(self, result_df: pd.DataFrame, output_path: Path) -> Path:
        csv_path = output_path.with_suffix(".csv")
        result_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        return csv_path
