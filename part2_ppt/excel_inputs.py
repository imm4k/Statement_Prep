from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

from openpyxl import load_workbook


@dataclass(frozen=True)
class GeneralConfig:
    statement_thru_date: datetime
    output_location: Path


def load_general_config(xlsx_path: Path, sheet_name: str, label_output_location: str, label_statement_thru: str) -> GeneralConfig:
    wb = load_workbook(filename=str(xlsx_path), data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Missing sheet '{sheet_name}' in setup workbook.")
    ws = wb[sheet_name]

    output_location = _find_label_value(ws, label_output_location)
    statement_thru = _find_label_value(ws, label_statement_thru)

    if output_location is None or str(output_location).strip() == "":
        raise ValueError("Could not locate Output Location value in General Config.")
    if statement_thru is None:
        raise ValueError("Could not locate Statement Thru Date value in General Config.")

    statement_thru_dt = _coerce_to_datetime(statement_thru)
    return GeneralConfig(statement_thru_date=statement_thru_dt, output_location=Path(str(output_location)))


def load_investors_from_run_config(xlsx_path: Path, sheet_name: str) -> List[str]:
    wb = load_workbook(filename=str(xlsx_path), data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Missing sheet '{sheet_name}' in setup workbook.")
    ws = wb[sheet_name]

    header_row_idx, investor_col_idx = _find_header_cell(ws, "Investor")
    if header_row_idx is None or investor_col_idx is None:
        raise ValueError("Could not find an 'Investor' header in Run Config.")

    investors: List[str] = []
    row_idx = header_row_idx + 1
    while True:
        val = ws.cell(row=row_idx, column=investor_col_idx).value
        if val is None or str(val).strip() == "":
            break
        investors.append(str(val).strip())
        row_idx += 1

    if not investors:
        raise ValueError("No investors found under the 'Investor' column in Run Config.")
    return investors


def _find_label_value(ws, label_text: str) -> Optional[object]:
    for row in ws.iter_rows(values_only=False):
        cell = row[0]
        if cell.value is None:
            continue
        if str(cell.value).strip() == label_text:
            return ws.cell(row=cell.row, column=cell.column + 1).value
    return None


def _find_header_cell(ws, header_text: str) -> Tuple[Optional[int], Optional[int]]:
    target = header_text.strip().lower()
    for r in range(1, min(ws.max_row, 200) + 1):
        for c in range(1, min(ws.max_column, 100) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            if str(v).strip().lower() == target:
                return r, c
    return None, None


def _coerce_to_datetime(value: object) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        s = value.strip()
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        raise ValueError(f"Unrecognized date string format for Statement Thru Date: {value}")
    raise ValueError(f"Unsupported Statement Thru Date value type: {type(value)}")
