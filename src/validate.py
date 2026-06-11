from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def sha256_12(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()[:12]


def normalize_cell(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def find_header_row(
    workbook_path: Path,
    required_columns: list[str],
    max_scan_rows: int = 30,
) -> dict[str, Any]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    required = {column.strip() for column in required_columns}

    try:
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row_index, row in enumerate(
                sheet.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True),
                start=1,
            ):
                values = [normalize_cell(cell) for cell in row]
                present = {value for value in values if value}
                missing = sorted(required - present)
                if not missing:
                    return {
                        "workbook_opened": True,
                        "sheet_name": sheet_name,
                        "header_row": row_index,
                        "required_columns_present": True,
                        "missing_required_columns": [],
                        "detected_columns": values,
                    }

        first_sheet = workbook[workbook.sheetnames[0]]
        first_row = next(first_sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        return {
            "workbook_opened": True,
            "sheet_name": None,
            "header_row": None,
            "required_columns_present": False,
            "missing_required_columns": sorted(required),
            "detected_columns_first_row": [normalize_cell(cell) for cell in first_row],
        }
    finally:
        workbook.close()


def validate_workbook(
    workbook_path: Path,
    required_columns: list[str],
    min_size_bytes: int,
) -> tuple[bool, str, dict[str, Any]]:
    metadata: dict[str, Any] = {
        "path": str(workbook_path),
        "size_bytes": workbook_path.stat().st_size,
        "min_size_bytes": min_size_bytes,
        "sha256_12": sha256_12(workbook_path),
    }

    if workbook_path.stat().st_size < min_size_bytes:
        return False, "artifact_too_small", metadata

    try:
        header_result = find_header_row(workbook_path, required_columns)
        metadata.update(header_result)
        if not header_result.get("required_columns_present"):
            return False, "required_columns_missing", metadata
        return True, "validation_passed", metadata
    except Exception as exc:
        metadata["exception_type"] = type(exc).__name__
        metadata["exception_message"] = str(exc)
        return False, "workbook_open_failed", metadata
