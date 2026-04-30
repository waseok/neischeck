from __future__ import annotations

from pathlib import Path
from typing import List

import openpyxl
import pandas as pd


class ExcelReader:
    SUPPORTED_SUFFIX = {".xlsx", ".xlsm"}

    def validate_path(self, path: Path) -> None:
        if path.suffix.lower() not in self.SUPPORTED_SUFFIX:
            raise ValueError("xlsx 또는 xlsm 파일만 지원합니다.")
        if not path.exists():
            raise FileNotFoundError(path)

    def list_sheets(self, path: Path) -> List[str]:
        self.validate_path(path)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_vba=True)
        names = wb.sheetnames
        wb.close()
        return names

    def read_preview(self, path: Path, sheet_name: str, header_row: int, nrows: int = 100) -> pd.DataFrame:
        self.validate_path(path)
        # xlsm도 매크로 실행 없이 셀 데이터만 읽는다.
        return pd.read_excel(path, sheet_name=sheet_name, header=header_row - 1, nrows=nrows, engine="openpyxl")

    def read_full(self, path: Path, sheet_name: str, header_row: int) -> pd.DataFrame:
        self.validate_path(path)
        return pd.read_excel(path, sheet_name=sheet_name, header=header_row - 1, engine="openpyxl")
