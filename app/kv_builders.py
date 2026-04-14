"""Builders that map KiotViet rows onto each target MTP template layout."""

from __future__ import annotations

from pathlib import Path

from .kv_excel import read_xls_rows, write_xls
from .kv_utils import (
    clean_text,
    make_unique_code,
    normalize_row,
    slugify,
    to_number,
    to_number_or_default,
)


def build_nganh_hang(loai_hang_values: list[str], existing_rows: list[list[object]]) -> list[list[object]]:
    # Build the ngành hàng sheet by keeping existing template rows and adding only
    # new unique names from the current product data.
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
        result.append(normalize_row(row, 3))

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
    # Build the nhóm hàng sheet and force column C to "MD" when it is blank, which
    # matches the preserved business rule in the original script.
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
        normalized = normalize_row(row, 4)
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


def build_san_pham(
    template_path: Path,
    output_path: Path,
    kiotviet_rows: list[dict[str, object]],
    merge_dvt: bool = False,
) -> int:
    # Merge KiotViet products into the product template.
    # Existing template rows are preserved, and new rows are deduped by product code
    # first, then by product name if the code is empty.
    headers, existing_rows = read_xls_rows(template_path)
    result = []
    existing_codes = set()

    for row in existing_rows:
        normalized = normalize_row(row, 82 if merge_dvt else 31)
        ma_san_pham = clean_text(normalized[1]) if len(normalized) > 1 else ""
        ten_san_pham = clean_text(normalized[2]) if len(normalized) > 2 else ""
        dedupe_key = ma_san_pham.casefold() or ten_san_pham.casefold()
        if dedupe_key:
            existing_codes.add(dedupe_key)
        if any(clean_text(cell) for cell in normalized):
            result.append(normalized)

    primary_rows = []
    variants_by_parent: dict[str, list[dict[str, object]]] = {}
    if merge_dvt:
        for item in kiotviet_rows:
            ma_dvt_co_ban = clean_text(item.get("ma_dvt_co_ban", ""))
            if ma_dvt_co_ban:
                variants_by_parent.setdefault(ma_dvt_co_ban, []).append(item)
            else:
                primary_rows.append(item)
    else:
        primary_rows = kiotviet_rows

    added = 0
    for item in primary_rows:
        ma_hang = clean_text(item["ma_hang"])
        ten_hang = clean_text(item["ten_hang"])
        if not ma_hang and not ten_hang:
            continue
        dedupe_key = ma_hang.casefold() or ten_hang.casefold()
        if dedupe_key in existing_codes:
            continue
        row = [""] * (82 if merge_dvt else 31)
        # These column positions match the MTP product template layout.
        row[0] = clean_text(item["nhom_hang"])
        row[1] = ma_hang
        row[2] = ten_hang
        row[4] = clean_text(item["don_vi_tinh"]) or "-"
        row[5] = to_number(item["gia_von"])
        row[6] = to_number(item["gia_ban"])
        row[7] = 0
        row[8] = 999999
        if item.get("ma_vach"):
            row[12] = clean_text(item.get("ma_vach", ""))

        if merge_dvt:
            variants = variants_by_parent.get(ma_hang, [])
            for i, variant in enumerate(variants[:20]):
                base = 22 + i * 3
                row[base] = clean_text(variant.get("don_vi_tinh", ""))
                row[base + 1] = to_number(variant.get("quy_doi", ""))
                row[base + 2] = to_number(variant.get("gia_ban", ""))

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


def build_ton_kho_dau_ky(
    template_path: Path,
    output_path: Path,
    kiotviet_rows: list[dict[str, object]],
    merge_dvt: bool = False,
    variant_codes: set[str] | None = None,
) -> int:
    # Build opening-stock rows, deduping by product code.
    headers, existing_rows = read_xls_rows(template_path)
    result = []
    existing_codes = set()

    for row in existing_rows:
        normalized = normalize_row(row, 3)
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
        if merge_dvt and variant_codes and ma_hang in variant_codes:
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


def build_kh_ncc(
    template_path: Path,
    output_path: Path,
    customer_rows: list[dict[str, object]],
    provider_rows: list[dict[str, object]],
) -> int:
    # Build the combined customer/provider master file.
    # Unlike product builders, this path appends incoming rows without deduping.
    headers, existing_rows = read_xls_rows(template_path)
    result = []

    for row in existing_rows:
        normalized = normalize_row(row, 13)
        if any(clean_text(cell) for cell in normalized):
            result.append(normalized)

    added = 0

    for item in customer_rows:
        row = [""] * 13
        # Customers are always marked as group KL and customer=TRUE.
        row[0] = clean_text(item["ma_khach_hang"])
        row[1] = clean_text(item["ten_khach_hang"])
        row[2] = clean_text(item["dia_chi"])
        row[4] = "KL"
        row[5] = "TRUE"
        row[7] = clean_text(item["dien_thoai"])
        result.append(row)
        added += 1

    for item in provider_rows:
        row = [""] * 13
        # Providers are always marked as group NCC and provider=TRUE.
        row[0] = clean_text(item["ma_nha_cung_cap"])
        row[1] = clean_text(item["ten_nha_cung_cap"])
        row[2] = clean_text(item["dia_chi"])
        row[4] = "NCC"
        row[6] = "TRUE"
        row[7] = clean_text(item["dien_thoai"])
        row[8] = clean_text(item["email"])
        result.append(row)
        added += 1

    write_xls(
        output_path,
        "Sheet1",
        headers or [
            "Mã KH-NCC",
            "Tên KH-NCC",
            "Địa chỉ",
            "Nhân viên phụ trách",
            "Nhóm KH-NCC",
            "Khách hàng(x)",
            "Nhà cung cấp(x)",
            "Điện thoại",
            "Email",
            "Ghi chú",
            "Điểm đầu kỳ",
            "Mã số thuế",
            "Ngày sinh nhật",
        ],
        result,
    )
    return added


def build_kh_cong_no_dau_ky(
    template_path: Path,
    output_path: Path,
    customer_rows: list[dict[str, object]],
    provider_rows: list[dict[str, object]],
) -> int:
    # Build opening receivable/payable balances for both customers and providers.
    headers, existing_rows = read_xls_rows(template_path)
    result = []

    for row in existing_rows:
        normalized = normalize_row(row, 4)
        if any(clean_text(cell) for cell in normalized):
            result.append(normalized)

    added = 0

    for item in customer_rows:
        row = [""] * 4
        row[0] = clean_text(item["ma_khach_hang"])
        row[2] = to_number_or_default(item["no_can_thu_hien_tai"], default=0)
        row[3] = 0
        result.append(row)
        added += 1

    for item in provider_rows:
        row = [""] * 4
        row[0] = clean_text(item["ma_nha_cung_cap"])
        row[2] = 0
        row[3] = to_number_or_default(item["no_can_tra_hien_tai"], default=0)
        result.append(row)
        added += 1

    write_xls(
        output_path,
        "Sheet1",
        headers or [
            "Mã KH - NCC(*)",
            "Tên KH - NCC",
            "Số tiền phải thu (*)",
            "Số tiền phải trả (*)",
        ],
        result,
    )
    return added
