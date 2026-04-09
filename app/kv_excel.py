"""Excel readers and writers for KiotViet inputs and MTP templates."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import xlrd
import xlwt
from openpyxl import load_workbook

from .kv_config import (
    CUSTOMER_HEADER_ALIASES,
    PRODUCT_HEADER_ALIASES,
    PROVIDER_HEADER_ALIASES,
)
from .kv_utils import clean_text, normalize_header


def resolve_columns(
    headers: list[str],
    aliases_map: dict[str, list[str]],
    source_label: str,
) -> dict[str, int]:
    # Build a "field name -> column number" mapping by matching normalized headers
    # against the alias lists above.
    normalized_headers = {
        normalize_header(header): idx + 1
        for idx, header in enumerate(headers)
        if clean_text(header)
    }
    resolved: dict[str, int] = {}
    missing: list[str] = []

    for field, aliases in aliases_map.items():
        match = None
        for alias in aliases:
            match = normalized_headers.get(normalize_header(alias))
            if match is not None:
                break
        if match is None:
            missing.append(f"{field} ({', '.join(aliases)})")
            continue
        resolved[field] = match

    if missing:
        raise ValueError(f"Thiếu cột bắt buộc trong file {source_label}: " + "; ".join(missing))

    return resolved


def read_xlsx_headers(path: Path) -> list[str]:
    # Read only the first row to detect the source type quickly.
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb.active
        return [clean_text(c.value) for c in ws[1]]
    finally:
        wb.close()


def read_mapped_xlsx_rows(
    path: Path,
    aliases_map: dict[str, list[str]],
    source_label: str,
    key_fields: tuple[str, ...],
) -> tuple[list[str], list[dict[str, object]]]:
    # Read a KiotViet sheet and return rows keyed by our internal field names.
    # Blank data rows are skipped based on the important identifying columns.
    wb = load_workbook(path, data_only=True)
    try:
        ws = wb.active
        headers = [clean_text(c.value) for c in ws[1]]
        columns = resolve_columns(headers, aliases_map, source_label)
        rows: list[dict[str, object]] = []

        for row_idx in range(2, ws.max_row + 1):
            row = {
                field: ws.cell(row_idx, col_idx).value
                for field, col_idx in columns.items()
            }
            if all(clean_text(row[field]) == "" for field in key_fields):
                continue
            rows.append(row)

        return headers, rows
    finally:
        wb.close()


def read_kiotviet_rows(path: Path) -> tuple[list[str], list[dict[str, object]]]:
    # Thin wrappers keep the per-source configuration close to the call site.
    return read_mapped_xlsx_rows(
        path,
        PRODUCT_HEADER_ALIASES,
        "KiotViet sản phẩm",
        ("ma_hang", "ten_hang"),
    )


def read_customer_rows(path: Path) -> tuple[list[str], list[dict[str, object]]]:
    return read_mapped_xlsx_rows(
        path,
        CUSTOMER_HEADER_ALIASES,
        "KiotViet khách hàng",
        ("ma_khach_hang", "ten_khach_hang"),
    )


def read_provider_rows(path: Path) -> tuple[list[str], list[dict[str, object]]]:
    return read_mapped_xlsx_rows(
        path,
        PROVIDER_HEADER_ALIASES,
        "KiotViet nhà cung cấp",
        ("ma_nha_cung_cap", "ten_nha_cung_cap"),
    )


def read_xls_rows(path: Path) -> tuple[list[str], list[list[object]]]:
    # Template files are .xls, so they are handled with xlrd instead of openpyxl.
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]
    headers = [clean_text(v) for v in rows[0]] if rows else []
    return headers, rows[1:] if len(rows) > 1 else []


def write_xls(
    path: Path,
    sheet_name: str,
    headers: list[str],
    rows: Iterable[Iterable[object]],
) -> None:
    # Write a simple .xls file from headers plus row data.
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)
    for c, value in enumerate(headers):
        ws.write(0, c, value)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            ws.write(r, c, value)
    wb.save(str(path))
