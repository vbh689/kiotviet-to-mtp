#!/usr/bin/env python3
# Venv setup: python3 -m venv venv ; source venv/bin/activate ; pip install xlrd xlwt openpyxl
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

import xlrd
import xlwt
from openpyxl import Workbook, load_workbook

KIOTVIET_HEADER_ALIASES = {
    "loai_hang": ["Loại hàng"],
    "nhom_hang": ["Nhóm hàng(3 Cấp)", "Nhóm hàng (3 Cấp)", "Nhóm hàng"],
    "ma_hang": ["Mã hàng"],
    "ten_hang": ["Tên hàng"],
    "gia_ban": ["Giá bán trước thuế", "Giá bán"],
    "gia_von": ["Giá vốn"],
    "ton_kho": ["Tồn kho"],
    "don_vi_tinh": ["ĐVT", "Đơn vị tính"],
}


def normalize_header(value) -> str:
    text = clean_text(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def resolve_kiotviet_columns(headers: list[str]) -> dict[str, int]:
    normalized_headers = {
        normalize_header(header): idx + 1
        for idx, header in enumerate(headers)
        if clean_text(header)
    }
    resolved: dict[str, int] = {}
    missing: list[str] = []

    for field, aliases in KIOTVIET_HEADER_ALIASES.items():
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
        raise ValueError(
            "Thiếu cột bắt buộc trong file KiotViet: " + "; ".join(missing)
        )

    return resolved


KNOWN_KIOTVIET_HEADERS = sorted(
    {alias for aliases in KIOTVIET_HEADER_ALIASES.values() for alias in aliases}
)


def clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        text = value.strip()
    else:
        text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = re.sub(r"[^A-Za-z0-9]+", "-", ascii_text).strip("-").upper()
    return ascii_text or "ITEM"


def make_unique_code(base: str, used: set[str], prefix: str = "") -> str:
    candidate = f"{prefix}{base}"
    if candidate not in used:
        used.add(candidate)
        return candidate
    index = 2
    while True:
        candidate = f"{prefix}{base}-{index}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        index += 1


def read_kiotviet_rows(path: Path) -> tuple[list[str], list[dict[str, object]]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [clean_text(c.value) for c in ws[1]]
    columns = resolve_kiotviet_columns(headers)
    rows: list[dict[str, object]] = []
    for row_idx in range(2, ws.max_row + 1):
        ma_hang = ws.cell(row_idx, columns["ma_hang"]).value
        ten_hang = ws.cell(row_idx, columns["ten_hang"]).value
        if clean_text(ma_hang) == "" and clean_text(ten_hang) == "":
            continue
        rows.append(
            {
                "loai_hang": ws.cell(row_idx, columns["loai_hang"]).value,
                "nhom_hang": ws.cell(row_idx, columns["nhom_hang"]).value,
                "ma_hang": ma_hang,
                "ten_hang": ten_hang,
                "gia_ban": ws.cell(row_idx, columns["gia_ban"]).value,
                "gia_von": ws.cell(row_idx, columns["gia_von"]).value,
                "ton_kho": ws.cell(row_idx, columns["ton_kho"]).value,
                "don_vi_tinh": ws.cell(row_idx, columns["don_vi_tinh"]).value,
            }
        )
    return headers, rows


def read_xls_rows(path: Path) -> tuple[list[str], list[list[object]]]:
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]
    headers = [clean_text(v) for v in rows[0]] if rows else []
    return headers, rows[1:] if len(rows) > 1 else []


def write_xls(path: Path, sheet_name: str, headers: list[str], rows: Iterable[Iterable[object]]) -> None:
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)
    for c, value in enumerate(headers):
        ws.write(0, c, value)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            ws.write(r, c, value)
    wb.save(str(path))


def to_number(value):
    if value is None or clean_text(value) == "":
        return None
    if isinstance(value, (int, float)):
        return value
    text = clean_text(value).replace(",", "")
    try:
        num = Decimal(text)
    except InvalidOperation:
        return value
    if num == num.to_integral_value():
        return int(num)
    return float(num)


def build_nganh_hang(loai_hang_values: list[str], existing_rows: list[list[object]]) -> list[list[object]]:
    result = []
    existing_names = set()
    used_codes = set()
    for row in existing_rows:
        name = clean_text(row[1]) if len(row) > 1 else ""
        code = clean_text(row[0]) if len(row) > 0 else ""
        if not any(clean_text(cell) for cell in row[:3]):
            continue
        if name:
            existing_names.add(name.casefold())
        if code:
            used_codes.add(code)
        result.append(list(row[:3]) + [""] * max(0, 3 - len(row)))

    for value in loai_hang_values:
        name = clean_text(value)
        if not name or name.casefold() in existing_names:
            continue
        suggested = "MD" if not used_codes else slugify(name)[:20]
        code = make_unique_code(suggested, used_codes)
        result.append([code, name, ""])
        existing_names.add(name.casefold())
    return result


def build_nhom_hang(nhom_hang_values: list[str], existing_rows: list[list[object]]) -> list[list[object]]:
    result = []
    existing_names = set()
    used_codes = set()
    for row in existing_rows:
        name = clean_text(row[1]) if len(row) > 1 else ""
        code = clean_text(row[0]) if len(row) > 0 else ""
        if not any(clean_text(cell) for cell in row[:4]):
            continue
        if name:
            existing_names.add(name.casefold())
        if code:
            used_codes.add(code)
        normalized = list(row[:4]) + [""] * max(0, 4 - len(row))
        if clean_text(normalized[2]) == "":
            normalized[2] = "MD"
        result.append(normalized)

    for value in nhom_hang_values:
        name = clean_text(value)
        if not name or name.casefold() in existing_names:
            continue
        code = make_unique_code(slugify(name)[:20], used_codes, prefix="NH-")
        result.append([code, name, "MD", ""])
        existing_names.add(name.casefold())
    return result


def build_san_pham(template_path: Path, output_path: Path, kiotviet_rows: list[dict[str, object]]) -> int:
    headers, existing_rows = read_xls_rows(template_path)
    result = []
    existing_codes = set()

    for row in existing_rows:
        normalized = list(row[:31]) + [""] * max(0, 31 - len(row))
        ma_san_pham = clean_text(normalized[1]) if len(normalized) > 1 else ""
        ten_san_pham = clean_text(normalized[2]) if len(normalized) > 2 else ""
        dedupe_key = ma_san_pham.casefold() or ten_san_pham.casefold()
        if dedupe_key:
            existing_codes.add(dedupe_key)
        if any(clean_text(cell) for cell in normalized):
            result.append(normalized)

    added = 0
    for item in kiotviet_rows:
        ma_hang = clean_text(item["ma_hang"])
        ten_hang = clean_text(item["ten_hang"])
        if not ma_hang and not ten_hang:
            continue
        dedupe_key = ma_hang.casefold() or ten_hang.casefold()
        if dedupe_key in existing_codes:
            continue
        row = [""] * 31
        row[0] = clean_text(item["nhom_hang"])
        row[1] = ma_hang
        row[2] = ten_hang
        row[4] = clean_text(item["don_vi_tinh"])
        row[5] = to_number(item["gia_von"])
        row[6] = to_number(item["gia_ban"])
        row[7] = 0
        row[8] = 999999
        result.append(row)
        existing_codes.add(dedupe_key)
        added += 1

    write_xls(
        output_path,
        "Sheet1",
        headers or [
            "Nhóm hàng hóa",
            "Mã sản phẩm",
            "Tên sản phẩm",
            "Tên viết tắt sản phẩm",
            "Đơn vị tính",
            "Giá nhập",
            "Giá bán",
            "Số lượng tồn tối thiểu",
            "Số lượng tồn tối đa",
            "Ghi chú",
        ],
        result,
    )
    return added


def build_ton_kho_dau_ky(template_path: Path, output_path: Path, kiotviet_rows: list[dict[str, object]]) -> int:
    headers, existing_rows = read_xls_rows(template_path)
    result = []
    existing_codes = set()

    for row in existing_rows:
        normalized = list(row[:3]) + [""] * max(0, 3 - len(row))
        ma_san_pham = clean_text(normalized[0])
        if ma_san_pham:
            existing_codes.add(ma_san_pham.casefold())
        if any(clean_text(cell) for cell in normalized):
            result.append(normalized)

    added = 0
    for item in kiotviet_rows:
        ma_hang = clean_text(item["ma_hang"])
        if not ma_hang:
            continue
        dedupe_key = ma_hang.casefold()
        if dedupe_key in existing_codes:
            continue
        result.append([
            ma_hang,
            to_number(item["ton_kho"]),
            to_number(item["gia_von"]),
        ])
        existing_codes.add(dedupe_key)
        added += 1

    write_xls(
        output_path,
        "Sheet1",
        headers or ["Mã sản phẩm", "Số lượng đầu kỳ", "Đơn giá đầu kỳ"],
        result,
    )
    return added


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Chuyển dữ liệu sản phẩm từ Excel KiotViet sang bộ file mẫu MTP."
    )
    parser.add_argument("--kiotviet", required=True, type=Path, help="File Excel KiotViet .xlsx")
    parser.add_argument(
        "--outdir",
        type=Path,
        default=Path(__file__).resolve().parent / "output",
        help="Thư mục xuất file (mặc định: ./output)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_dir = Path(__file__).resolve().parent
    templates_dir = project_dir / "templates"
    mtp_nganh = templates_dir / "MTP-NganhHang-LoaiHang.xls"
    mtp_nhom = templates_dir / "MTP-NhomHang.xls"
    mtp_sanpham = templates_dir / "MTP-SanPham.xls"
    ton_kho_dau_ky = templates_dir / "Mau-TonKhoDauKy.xls"

    missing_templates = [
        path.name
        for path in (mtp_nganh, mtp_nhom, mtp_sanpham, ton_kho_dau_ky)
        if not path.exists()
    ]
    if missing_templates:
        print(
            "Thiếu file template trong thư mục templates/: "
            + ", ".join(missing_templates),
            file=sys.stderr,
        )
        return 1

    args.outdir.mkdir(parents=True, exist_ok=True)

    try:
        _, kv_rows = read_kiotviet_rows(args.kiotviet)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        print(
            "Các tiêu đề đang hỗ trợ: " + ", ".join(KNOWN_KIOTVIET_HEADERS),
            file=sys.stderr,
        )
        return 1
    loai_hang_values = [clean_text(r["loai_hang"]) for r in kv_rows]
    nhom_hang_values = [clean_text(r["nhom_hang"]) for r in kv_rows]

    nganh_headers, nganh_existing = read_xls_rows(mtp_nganh)
    nhom_headers, nhom_existing = read_xls_rows(mtp_nhom)

    nganh_rows = build_nganh_hang(loai_hang_values, nganh_existing)
    nhom_rows = build_nhom_hang(nhom_hang_values, nhom_existing)

    out_nganh = args.outdir / "MTP-NganhHang-LoaiHang.xls"
    out_nhom = args.outdir / "MTP-NhomHang.xls"
    out_sanpham = args.outdir / "MTP-SanPham.xls"
    out_ton_kho = args.outdir / "Mau-TonKhoDauKy.xls"

    write_xls(out_nganh, "Sheet2", nganh_headers or ["Mã ngành", "Tên ngành", "Ghi chú"], nganh_rows)
    write_xls(out_nhom, "Sheet2", nhom_headers or ["Mã nhóm", "Tên nhóm", "Ngành hàng", "Ghi chú"], nhom_rows)
    san_pham_added = build_san_pham(mtp_sanpham, out_sanpham, kv_rows)
    ton_kho_added = build_ton_kho_dau_ky(ton_kho_dau_ky, out_ton_kho, kv_rows)

    print(f"Đã xuất: {out_nganh}")
    print(f"Đã xuất: {out_nhom}")
    print(f"Đã xuất: {out_sanpham}")
    print(f"Đã xuất: {out_ton_kho}")
    print(f"Số dòng sản phẩm thêm mới: {san_pham_added}")
    print(f"Số dòng tồn kho đầu kỳ thêm mới: {ton_kho_added}")
    print(f"Số ngành hàng duy nhất: {len(nganh_rows)}")
    print(f"Số nhóm hàng duy nhất: {len(nhom_rows)}")
    print("Ghi chú: cột C của MTP-NhomHang được đặt cố định là 'MD'; cột H/I của MTP-SanPham lần lượt là 0 và 999999.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
