"""CLI orchestration for reading KiotViet exports and writing MTP outputs."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .kv_builders import (
    build_kh_cong_no_dau_ky,
    build_kh_ncc,
    build_nganh_hang,
    build_nhom_hang,
    build_san_pham,
    build_ton_kho_dau_ky,
)
from .kv_config import (
    CUSTOMER_HEADER_ALIASES,
    KNOWN_CUSTOMER_HEADERS,
    KNOWN_PRODUCT_HEADERS,
    KNOWN_PROVIDER_HEADERS,
    PRODUCT_HEADER_ALIASES,
    PROVIDER_HEADER_ALIASES,
    SOURCE_PREFIXES,
    TEMPLATE_CANDIDATES,
)
from .kv_excel import (
    read_customer_rows,
    read_kiotviet_rows,
    read_provider_rows,
    read_xls_rows,
    read_xlsx_headers,
    write_xls,
)
from .kv_mapping import ColumnMappings
from .kv_utils import clean_text, flatten_aliases, normalize_header


def get_project_dir() -> Path:
    # This module lives under app/, but templates/ and the default output folder
    # are still rooted at the repository level.
    return Path(__file__).resolve().parent.parent


def get_templates_dir() -> Path:
    if getattr(sys, 'frozen', False):
        # PyInstaller bundled execution
        return Path(sys._MEIPASS) / "templates"
    return get_project_dir() / "templates"


def get_default_outdir() -> Path:
    if getattr(sys, 'frozen', False):
        # PyInstaller bundled: write next to where it's executed
        exe_path = Path(sys.executable).resolve()
        # On macOS inside a windowed .app bundle, sys.executable is 
        # path/to/MyApp.app/Contents/MacOS/MyApp
        if sys.platform == "darwin" and exe_path.parent.name == "MacOS" and exe_path.parent.parent.name == "Contents":
            return exe_path.parent.parent.parent.parent / "output"
        return exe_path.parent / "output"
    return get_project_dir() / "output"


def detect_source_type(path: Path, headers: list[str]) -> str:
    # Detection order matters:
    # 1. filename prefix (fast and explicit)
    # 2. fallback to key headers when the filename is less reliable
    normalized_name = normalize_header(path.stem)
    for source_type, prefix in SOURCE_PREFIXES.items():
        if normalized_name.startswith(normalize_header(prefix)):
            return source_type

    normalized_headers = {normalize_header(header) for header in headers if clean_text(header)}
    if {
        normalize_header("Mã hàng"),
        normalize_header("Tên hàng"),
    }.issubset(normalized_headers):
        return "product"
    if {
        normalize_header("Mã khách hàng"),
        normalize_header("Tên khách hàng"),
    }.issubset(normalized_headers):
        return "customer"
    if {
        normalize_header("Mã nhà cung cấp"),
        normalize_header("Tên nhà cung cấp"),
    }.issubset(normalized_headers):
        return "provider"

    raise ValueError(
        f"Không nhận diện được loại file: {path.name}. "
        "Tên file nên bắt đầu bằng DanhSachSanPham, DanhSachKhachHang hoặc DanhSachNhaCungCap."
    )


def resolve_template_path(templates_dir: Path, candidates: list[str]) -> Path | None:
    # Return the first template filename that exists on disk.
    for name in candidates:
        path = templates_dir / name
        if path.exists():
            return path
    return None


def parse_args() -> argparse.Namespace:
    # CLI stays intentionally small: input KiotViet files plus an optional output folder.
    parser = argparse.ArgumentParser(
        description="Chuyển dữ liệu sản phẩm, khách hàng và nhà cung cấp từ Excel KiotViet sang bộ file mẫu MTP."
    )
    parser.add_argument(
        "--kiotviet",
        required=False,
        nargs="+",
        type=Path,
        help="Một hoặc nhiều file Excel KiotViet .xlsx (Bỏ trống để mở chế độ giao diện GUI)",
    )
    parser.add_argument(
        "--outdir",
        type=Path,
        default=get_default_outdir(),
        help="Thư mục xuất file (mặc định: ./output)",
    )
    return parser.parse_args()


def convert_kiotviet_files(
    source_paths: list[Path],
    outdir: Path,
    column_mappings: ColumnMappings | None = None,
) -> int:
    # Shared conversion flow:
    # 1. read and classify each source file
    # 2. resolve the needed templates
    # 3. generate only the output groups that have input data
    templates_dir = get_templates_dir()

    product_rows: list[dict[str, object]] = []
    customer_rows: list[dict[str, object]] = []
    provider_rows: list[dict[str, object]] = []
    source_counts = {"product": 0, "customer": 0, "provider": 0}

    # A single run can mix products, customers, and providers.
    for source_path in source_paths:
        source_path = Path(source_path)
        if not source_path.exists():
            print(f"Không tìm thấy file: {source_path}", file=sys.stderr)
            return 1

        source_type = "unknown"
        try:
            headers = read_xlsx_headers(source_path)
            source_type = detect_source_type(source_path, headers)
            selected_mapping = None
            if column_mappings is not None:
                selected_mapping = column_mappings.get(source_type)
            if source_type == "product":
                _, rows = read_kiotviet_rows(source_path, selected_mapping)
                product_rows.extend(rows)
            elif source_type == "customer":
                _, rows = read_customer_rows(source_path, selected_mapping)
                customer_rows.extend(rows)
            else:
                _, rows = read_provider_rows(source_path, selected_mapping)
                provider_rows.extend(rows)
            source_counts[source_type] += 1
        except ValueError as exc:
            # When parsing fails, also print the currently supported headers to help
            # the user compare their export file against the expected format.
            print(str(exc), file=sys.stderr)
            if source_type == "product":
                supported_headers = KNOWN_PRODUCT_HEADERS
            elif source_type == "customer":
                supported_headers = KNOWN_CUSTOMER_HEADERS
            elif source_type == "provider":
                supported_headers = KNOWN_PROVIDER_HEADERS
            else:
                supported_headers = flatten_aliases(
                    PRODUCT_HEADER_ALIASES,
                    CUSTOMER_HEADER_ALIASES,
                    PROVIDER_HEADER_ALIASES,
                )
            print(
                "Các tiêu đề đang hỗ trợ: " + ", ".join(supported_headers),
                file=sys.stderr,
            )
            return 1

    needs_product_outputs = bool(product_rows)
    needs_partner_outputs = bool(customer_rows or provider_rows)

    missing_templates: list[str] = []
    resolved_templates: dict[str, Path] = {}

    # Only require the templates that are needed for the data types present in this run.
    for key, candidates in TEMPLATE_CANDIDATES.items():
        if key.startswith("product_") and not needs_product_outputs:
            continue
        if key.startswith("kh_") and not needs_partner_outputs:
            continue
        path = resolve_template_path(templates_dir, candidates)
        if path is None:
            missing_templates.append("/".join(candidates))
            continue
        resolved_templates[key] = path

    if missing_templates:
        print(
            "Thiếu file template trong thư mục templates/: "
            + ", ".join(missing_templates),
            file=sys.stderr,
        )
        return 1

    outdir.mkdir(parents=True, exist_ok=True)

    if needs_product_outputs:
        # Product data produces four separate MTP files.
        mtp_nganh = resolved_templates["product_nganh"]
        mtp_nhom = resolved_templates["product_nhom"]
        mtp_sanpham = resolved_templates["product_sanpham"]
        ton_kho_dau_ky = resolved_templates["product_tonkho"]

        loai_hang_values = [clean_text(r["loai_hang"]) for r in product_rows]
        nhom_hang_values = [clean_text(r["nhom_hang"]) for r in product_rows]
        nganh_headers, nganh_existing = read_xls_rows(mtp_nganh)
        nhom_headers, nhom_existing = read_xls_rows(mtp_nhom)
        nganh_rows = build_nganh_hang(loai_hang_values, nganh_existing)
        nhom_rows = build_nhom_hang(nhom_hang_values, nhom_existing)

        out_nganh = outdir / mtp_nganh.name
        out_nhom = outdir / mtp_nhom.name
        out_sanpham = outdir / mtp_sanpham.name
        out_ton_kho = outdir / ton_kho_dau_ky.name

        write_xls(
            out_nganh,
            "Sheet2",
            nganh_headers or ["Mã ngành", "Tên ngành", "Ghi chú"],
            nganh_rows,
        )
        write_xls(
            out_nhom,
            "Sheet2",
            nhom_headers or ["Mã nhóm", "Tên nhóm", "Ngành hàng", "Ghi chú"],
            nhom_rows,
        )
        san_pham_added = build_san_pham(mtp_sanpham, out_sanpham, product_rows)
        ton_kho_added = build_ton_kho_dau_ky(ton_kho_dau_ky, out_ton_kho, product_rows)

        print(f"Đã xuất: {out_nganh}")
        print(f"Đã xuất: {out_nhom}")
        print(f"Đã xuất: {out_sanpham}")
        print(f"Đã xuất: {out_ton_kho}")
        print(f"Số dòng sản phẩm thêm mới: {san_pham_added}")
        print(f"Số dòng tồn kho đầu kỳ thêm mới: {ton_kho_added}")
        print(f"Số ngành hàng duy nhất: {len(nganh_rows)}")
        print(f"Số nhóm hàng duy nhất: {len(nhom_rows)}")
        print(
            "Ghi chú: cột C của MTP_SP-NhomHang được đặt cố định là 'MD'; "
            "cột H/I của MTP_SP-SanPham lần lượt là 0 và 999999."
        )

    if needs_partner_outputs:
        # Customer and provider data share the same KH/NCC templates.
        kh_ncc_template = resolved_templates["kh_ncc"]
        kh_congno_template = resolved_templates["kh_congno"]
        out_kh_ncc = outdir / kh_ncc_template.name
        out_kh_congno = outdir / kh_congno_template.name
        kh_ncc_added = build_kh_ncc(kh_ncc_template, out_kh_ncc, customer_rows, provider_rows)
        kh_congno_added = build_kh_cong_no_dau_ky(
            kh_congno_template,
            out_kh_congno,
            customer_rows,
            provider_rows,
        )

        print(f"Đã xuất: {out_kh_ncc}")
        print(f"Đã xuất: {out_kh_congno}")
        print(f"Số dòng KH/NCC thêm mới: {kh_ncc_added}")
        print(f"Số dòng công nợ đầu kỳ KH/NCC thêm mới: {kh_congno_added}")
        print(
            "Ghi chú: KH mặc định dùng nhóm 'KL' và cột Khách hàng(x) = TRUE; "
            "NCC dùng nhóm 'NCC' và cột Nhà cung cấp(x) = TRUE."
        )

    print(
        "Đã xử lý file nguồn: "
        f"{source_counts['product']} sản phẩm, "
        f"{source_counts['customer']} khách hàng, "
        f"{source_counts['provider']} nhà cung cấp."
    )
    return 0


def main() -> int:
    args = parse_args()

    # Fallback to GUI mode if no files were passed via CLI.
    if not args.kiotviet:
        try:
            from .gui import run_gui
            return run_gui()
        except ImportError as e:
            print(f"Lỗi: Không thể khởi chạy giao diện do thiếu thư viện PyQt6 ({e}).", file=sys.stderr)
            print("Chạy `pip install PyQt6` để cài đặt giao diện, hoặc truyền đối số `--kiotviet` để chạy giao diện dòng lệnh.", file=sys.stderr)
            return 1

    return convert_kiotviet_files(args.kiotviet, args.outdir)
