from __future__ import annotations

import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from app.kv_builders import build_kh_ncc, build_san_pham, build_ton_kho_dau_ky
from app.kv_config import PRODUCT_HEADER_ALIASES
from app.kv_excel import read_kiotviet_rows, read_xls_rows, resolve_columns, write_xls
from app.kv_mapping import ColumnMappings
from app.kv_runner import convert_kiotviet_files, detect_source_type, resolve_template_path


def write_xlsx(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def find_row(rows: list[list[object]], column_idx: int, value: str) -> list[object]:
    for row in rows:
        if len(row) > column_idx and row[column_idx] == value:
            return row
    raise AssertionError(f"Could not find row with column {column_idx}={value!r}")


def run_conversion_quietly(*args, **kwargs) -> int:
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        return convert_kiotviet_files(*args, **kwargs)


class RefactorBehaviorTests(unittest.TestCase):
    def test_detect_source_type_prefers_filename_prefix(self):
        headers = ["Mã khách hàng", "Tên khách hàng"]
        source_type = detect_source_type(Path("DanhSachSanPham_fake.xlsx"), headers)
        self.assertEqual(source_type, "product")

    def test_detect_source_type_falls_back_to_headers(self):
        headers = ["Mã nhà cung cấp", "Tên nhà cung cấp"]
        source_type = detect_source_type(Path("some-file.xlsx"), headers)
        self.assertEqual(source_type, "provider")

    def test_resolve_columns_and_missing_column_message(self):
        headers = ["Loại hàng", "Nhóm hàng", "Mã hàng", "Tên hàng", "Giá bán", "Giá vốn", "Tồn kho", "ĐVT"]
        columns = resolve_columns(headers, PRODUCT_HEADER_ALIASES, "KiotViet sản phẩm")
        self.assertEqual(columns["ma_hang"], 3)
        self.assertEqual(columns["don_vi_tinh"], 8)

        with self.assertRaises(ValueError) as ctx:
            resolve_columns(["Mã hàng"], PRODUCT_HEADER_ALIASES, "KiotViet sản phẩm")
        self.assertIn("Thiếu cột bắt buộc", str(ctx.exception))
        self.assertIn("ten_hang", str(ctx.exception))

    def test_read_product_rows_with_explicit_reordered_mapping(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "DanhSachSanPham_custom.xlsx"
            write_xlsx(
                path,
                ["Tên custom", "Mã custom", "Tồn custom", "Vốn custom", "Nhóm custom", "Loại custom", "Bán custom", "ĐVT custom"],
                [["Tên chọn", "SP-MAP-001", 12, 700, "Nhóm chọn", "Loại chọn", 900, "Hộp"]],
            )

            mapping = {
                "loai_hang": 6,
                "nhom_hang": 5,
                "ma_hang": 2,
                "ten_hang": 1,
                "gia_ban": 7,
                "gia_von": 4,
                "ton_kho": 3,
                "don_vi_tinh": 8,
            }

            _, rows = read_kiotviet_rows(path, mapping)

            self.assertEqual(rows[0]["ma_hang"], "SP-MAP-001")
            self.assertEqual(rows[0]["ten_hang"], "Tên chọn")
            self.assertEqual(rows[0]["nhom_hang"], "Nhóm chọn")
            self.assertEqual(rows[0]["gia_von"], 700)

    def test_explicit_mapping_reports_missing_and_invalid_columns(self):
        headers = ["Mã custom"]
        with self.assertRaises(ValueError) as missing_ctx:
            resolve_columns(
                headers,
                PRODUCT_HEADER_ALIASES,
                "KiotViet sản phẩm",
                {"ma_hang": 1},
            )
        self.assertIn("Thiếu mapping bắt buộc", str(missing_ctx.exception))
        self.assertIn("ten_hang", str(missing_ctx.exception))

        full_mapping = {
            "loai_hang": 1,
            "nhom_hang": 1,
            "ma_hang": 2,
            "ten_hang": 1,
            "gia_ban": 1,
            "gia_von": 1,
            "ton_kho": 1,
            "don_vi_tinh": 1,
        }
        with self.assertRaises(ValueError) as invalid_ctx:
            resolve_columns(headers, PRODUCT_HEADER_ALIASES, "KiotViet sản phẩm", full_mapping)
        self.assertIn("Mapping cột không hợp lệ", str(invalid_ctx.exception))
        self.assertIn("ma_hang", str(invalid_ctx.exception))

    def test_build_san_pham_uses_default_unit_dash(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            template = tmp / "MTP_SP-SanPham.xls"
            output = tmp / "out.xls"

            write_xls(template, "Sheet1", ["Nhóm hàng hóa"], [])
            added = build_san_pham(
                template,
                output,
                [
                    {
                        "nhom_hang": "Nhóm A",
                        "ma_hang": "SP001",
                        "ten_hang": "Sản phẩm 1",
                        "don_vi_tinh": "",
                        "gia_von": "1000",
                        "gia_ban": "1500",
                    }
                ],
            )

            headers, rows = read_xls_rows(output)
            self.assertEqual(added, 1)
            self.assertEqual(headers[0], "Nhóm hàng hóa")
            self.assertEqual(rows[0][4], "-")
            self.assertEqual(rows[0][7], 0.0)
            self.assertEqual(rows[0][8], 999999.0)

    def test_product_and_opening_stock_dedupe(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            product_template = tmp / "MTP_SP-SanPham.xls"
            product_output = tmp / "product_out.xls"
            stock_template = tmp / "MTP_SP-TonKhoDauKy.xls"
            stock_output = tmp / "stock_out.xls"

            write_xls(product_template, "Sheet1", ["Nhóm hàng hóa"], [["", "SP001", "Đã có"]])
            write_xls(stock_template, "Sheet1", ["Mã sản phẩm"], [["SP001", 2, 10]])

            source_rows = [
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "SP001",
                    "ten_hang": "Sản phẩm trùng template",
                    "don_vi_tinh": "Hộp",
                    "gia_von": 1,
                    "gia_ban": 2,
                    "ton_kho": 3,
                },
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "SP002",
                    "ten_hang": "Sản phẩm mới",
                    "don_vi_tinh": "Hộp",
                    "gia_von": 4,
                    "gia_ban": 5,
                    "ton_kho": 6,
                },
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "SP002",
                    "ten_hang": "Sản phẩm mới trùng input",
                    "don_vi_tinh": "Hộp",
                    "gia_von": 7,
                    "gia_ban": 8,
                    "ton_kho": 9,
                },
            ]

            added_products = build_san_pham(product_template, product_output, source_rows)
            added_stock = build_ton_kho_dau_ky(stock_template, stock_output, source_rows)

            _, product_rows = read_xls_rows(product_output)
            _, stock_rows = read_xls_rows(stock_output)

            self.assertEqual(added_products, 1)
            self.assertEqual(added_stock, 1)
            self.assertEqual(len(product_rows), 2)
            self.assertEqual(len(stock_rows), 2)
            self.assertEqual(product_rows[1][1], "SP002")
            self.assertEqual(stock_rows[1][0], "SP002")

    def test_build_kh_ncc_sets_default_flags_and_groups(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            template = tmp / "MTP_KH-NCC.xls"
            output = tmp / "out.xls"

            write_xls(template, "Sheet1", ["Mã KH-NCC"], [])
            added = build_kh_ncc(
                template,
                output,
                [
                    {
                        "ma_khach_hang": "KH01",
                        "ten_khach_hang": "Khách A",
                        "dia_chi": "HCM",
                        "dien_thoai": "0901",
                    }
                ],
                [
                    {
                        "ma_nha_cung_cap": "NCC01",
                        "ten_nha_cung_cap": "NCC A",
                        "dia_chi": "HN",
                        "dien_thoai": "0902",
                        "email": "ncc@example.com",
                    }
                ],
            )

            _, rows = read_xls_rows(output)
            self.assertEqual(added, 2)
            self.assertEqual(rows[0][4], "KL")
            self.assertEqual(rows[0][5], "TRUE")
            self.assertEqual(rows[1][4], "NCC")
            self.assertEqual(rows[1][6], "TRUE")
            self.assertEqual(rows[1][8], "ncc@example.com")

    def test_resolve_template_path_accepts_legacy_filename(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            legacy = tmp / "MTP-SanPham.xls"
            legacy.touch()

            resolved = resolve_template_path(
                tmp,
                ["MTP_SP-SanPham.xls", "MTP-SanPham.xls"],
            )
            self.assertEqual(resolved, legacy)

    def test_conversion_product_uses_selected_columns(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            source = tmp / "DanhSachSanPham_custom.xlsx"
            outdir = tmp / "output"
            write_xlsx(
                source,
                ["Tên custom", "Mã custom", "Tồn custom", "Vốn custom", "Nhóm custom", "Loại custom", "Bán custom", "ĐVT custom"],
                [["Tên mapping sản phẩm", "SP-MAP-CONVERT-001", 23, 1100, "Nhóm mapping", "Loại mapping", 1500, "Thùng"]],
            )
            mappings: ColumnMappings = {
                "product": {
                    "loai_hang": 6,
                    "nhom_hang": 5,
                    "ma_hang": 2,
                    "ten_hang": 1,
                    "gia_ban": 7,
                    "gia_von": 4,
                    "ton_kho": 3,
                    "don_vi_tinh": 8,
                }
            }

            code = run_conversion_quietly([source], outdir, mappings)

            self.assertEqual(code, 0)
            _, product_rows = read_xls_rows(outdir / "MTP_SP-SanPham.xls")
            _, stock_rows = read_xls_rows(outdir / "MTP_SP-TonKhoDauKy.xls")
            product = find_row(product_rows, 1, "SP-MAP-CONVERT-001")
            stock = find_row(stock_rows, 0, "SP-MAP-CONVERT-001")
            self.assertEqual(product[0], "Nhóm mapping")
            self.assertEqual(product[2], "Tên mapping sản phẩm")
            self.assertEqual(product[4], "Thùng")
            self.assertEqual(product[5], 1100.0)
            self.assertEqual(product[6], 1500.0)
            self.assertEqual(stock[1], 23.0)
            self.assertEqual(stock[2], 1100.0)

    def test_conversion_customer_provider_uses_selected_columns(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            customer_source = tmp / "DanhSachKhachHang_custom.xlsx"
            provider_source = tmp / "DanhSachNhaCungCap_custom.xlsx"
            outdir = tmp / "output"
            write_xlsx(
                customer_source,
                ["Tên custom", "Mã custom", "Nợ custom", "Địa chỉ custom", "Điện thoại custom"],
                [["Khách mapping", "KH-MAP-001", 3200, "HCM mapping", "0909000001"]],
            )
            write_xlsx(
                provider_source,
                ["Tên custom", "Mã custom", "Email custom", "Nợ custom", "Địa chỉ custom", "Điện thoại custom"],
                [["NCC mapping", "NCC-MAP-001", "ncc-map@example.com", 4300, "HN mapping", "0909000002"]],
            )
            mappings: ColumnMappings = {
                "customer": {
                    "ma_khach_hang": 2,
                    "ten_khach_hang": 1,
                    "dien_thoai": 5,
                    "dia_chi": 4,
                    "no_can_thu_hien_tai": 3,
                },
                "provider": {
                    "ma_nha_cung_cap": 2,
                    "ten_nha_cung_cap": 1,
                    "email": 3,
                    "dien_thoai": 6,
                    "dia_chi": 5,
                    "no_can_tra_hien_tai": 4,
                },
            }

            code = run_conversion_quietly([customer_source, provider_source], outdir, mappings)

            self.assertEqual(code, 0)
            _, partner_rows = read_xls_rows(outdir / "MTP_KH-NCC.xls")
            _, debt_rows = read_xls_rows(outdir / "MTP_KH-CongNoDauKy.xls")
            customer = find_row(partner_rows, 0, "KH-MAP-001")
            provider = find_row(partner_rows, 0, "NCC-MAP-001")
            customer_debt = find_row(debt_rows, 0, "KH-MAP-001")
            provider_debt = find_row(debt_rows, 0, "NCC-MAP-001")
            self.assertEqual(customer[1], "Khách mapping")
            self.assertEqual(customer[2], "HCM mapping")
            self.assertEqual(customer[4], "KL")
            self.assertEqual(customer[5], "TRUE")
            self.assertEqual(customer[7], "0909000001")
            self.assertEqual(provider[1], "NCC mapping")
            self.assertEqual(provider[2], "HN mapping")
            self.assertEqual(provider[4], "NCC")
            self.assertEqual(provider[6], "TRUE")
            self.assertEqual(provider[7], "0909000002")
            self.assertEqual(provider[8], "ncc-map@example.com")
            self.assertEqual(customer_debt[2], 3200.0)
            self.assertEqual(customer_debt[3], 0.0)
            self.assertEqual(provider_debt[2], 0.0)
            self.assertEqual(provider_debt[3], 4300.0)

    def test_conversion_without_mapping_still_uses_default_aliases(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            source = tmp / "DanhSachSanPham_alias.xlsx"
            outdir = tmp / "output"
            write_xlsx(
                source,
                ["Loại hàng", "Nhóm hàng", "Mã hàng", "Tên hàng", "Giá bán", "Giá vốn", "Tồn kho", "ĐVT"],
                [["Loại alias", "Nhóm alias", "SP-ALIAS-001", "Tên alias", 2000, 1400, 5, "Cái"]],
            )

            code = run_conversion_quietly([source], outdir)

            self.assertEqual(code, 0)
            _, rows = read_xls_rows(outdir / "MTP_SP-SanPham.xls")
            product = find_row(rows, 1, "SP-ALIAS-001")
            self.assertEqual(product[2], "Tên alias")
            self.assertEqual(product[4], "Cái")

    def test_cli_wrapper_keeps_original_entrypoint(self):
        repo_root = Path(__file__).resolve().parents[1]
        result = subprocess.run(
            [sys.executable, "kiotviet_to_mtp.py", "--help"],
            cwd=repo_root,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(result.returncode, 0)
        self.assertIn("--kiotviet", result.stdout)

    def test_build_san_pham_merge_dvt_basic(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            template = tmp / "MTP_SP-SanPham.xls"
            output = tmp / "out.xls"

            write_xls(template, "Sheet1", ["Nhóm hàng hóa"], [])
            # 1 primary + 2 variants
            source_rows = [
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "PVN3220",
                    "ten_hang": "Bài Double K Xịn",
                    "don_vi_tinh": "Bộ",
                    "gia_von": "1000",
                    "gia_ban": "1500",
                    "ma_dvt_co_ban": "",
                    "quy_doi": "",
                },
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "SP001205",
                    "ten_hang": "Bài Double K Xịn",
                    "don_vi_tinh": "Cây",
                    "gia_von": "10000",
                    "gia_ban": "15000",
                    "ma_dvt_co_ban": "PVN3220",
                    "quy_doi": "10",
                },
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "SP001206",
                    "ten_hang": "Bài Double K Xịn",
                    "don_vi_tinh": "Thùng",
                    "gia_von": "50000",
                    "gia_ban": "75000",
                    "ma_dvt_co_ban": "PVN3220",
                    "quy_doi": "50",
                },
            ]

            added = build_san_pham(template, output, source_rows, merge_dvt=True)

            headers, rows = read_xls_rows(output)
            self.assertEqual(added, 1)  # Only 1 output row
            self.assertEqual(len(rows), 1)
            # Slot 01: W(22)=Cây, X(23)=10, Y(24)=15000
            self.assertEqual(rows[0][22], "Cây")
            self.assertEqual(rows[0][23], 10.0)
            self.assertEqual(rows[0][24], 15000.0)
            # Slot 02: Z(25)=Thùng, AA(26)=50, AB(27)=75000
            self.assertEqual(rows[0][25], "Thùng")
            self.assertEqual(rows[0][26], 50.0)
            self.assertEqual(rows[0][27], 75000.0)

    def test_build_san_pham_merge_dvt_no_variants(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            template = tmp / "MTP_SP-SanPham.xls"
            output = tmp / "out.xls"

            write_xls(template, "Sheet1", ["Nhóm hàng hóa"], [])
            source_rows = [
                {
                    "nhom_hang": "Nhóm B",
                    "ma_hang": "SP001",
                    "ten_hang": "Sản phẩm 1",
                    "don_vi_tinh": "Cái",
                    "gia_von": "100",
                    "gia_ban": "200",
                    "ma_dvt_co_ban": "",
                    "quy_doi": "",
                }
            ]

            added = build_san_pham(template, output, source_rows, merge_dvt=True)

            headers, rows = read_xls_rows(output)
            self.assertEqual(added, 1)
            self.assertEqual(len(rows), 1)
            self.assertTrue(len(rows[0]) >= 9)  # xlrd may strip trailing empty cells
            self.assertEqual(rows[0][1], "SP001")

    def test_build_san_pham_merge_dvt_off(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            template = tmp / "MTP_SP-SanPham.xls"
            output = tmp / "out.xls"

            write_xls(template, "Sheet1", ["Nhóm hàng hóa"], [])
            source_rows = [
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "PVN3220",
                    "ten_hang": "Bài Double K Xịn",
                    "don_vi_tinh": "Bộ",
                    "gia_von": "1000",
                    "gia_ban": "1500",
                    "ma_dvt_co_ban": "",
                    "quy_doi": "",
                },
                {
                    "nhom_hang": "Nhóm A",
                    "ma_hang": "SP001205",
                    "ten_hang": "Bài Double K Xịn",
                    "don_vi_tinh": "Cây",
                    "gia_von": "10000",
                    "gia_ban": "15000",
                    "ma_dvt_co_ban": "PVN3220",
                    "quy_doi": "10",
                },
            ]

            added = build_san_pham(template, output, source_rows, merge_dvt=False)

            headers, rows = read_xls_rows(output)
            self.assertEqual(added, 2)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0][1], "PVN3220")
            self.assertEqual(rows[1][1], "SP001205")

    def test_ton_kho_excludes_variants_when_merge_dvt(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            template = tmp / "MTP_SP-TonKhoDauKy.xls"
            output = tmp / "out.xls"

            write_xls(template, "Sheet1", ["Mã sản phẩm"], [])
            source_rows = [
                {"ma_hang": "MAIN", "ton_kho": "10", "gia_von": "100"},
                {"ma_hang": "VARIANT1", "ton_kho": "5", "gia_von": "50"},
            ]

            added = build_ton_kho_dau_ky(template, output, source_rows, merge_dvt=True, variant_codes={"VARIANT1"})

            headers, rows = read_xls_rows(output)
            self.assertEqual(added, 1)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0][0], "MAIN")

    def test_optional_columns_missing_no_error(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "DanhSachSanPham_missing_optional.xlsx"
            write_xlsx(
                path,
                ["Loại hàng", "Nhóm hàng", "Mã hàng", "Tên hàng", "Giá bán", "Giá vốn", "Tồn kho", "ĐVT"],
                [["Loại 1", "Nhóm 1", "SP01", "Tên 1", 200, 100, 10, "Cái"]],
            )

            _, rows = read_kiotviet_rows(path)
            self.assertEqual(rows[0]["ma_hang"], "SP01")
            self.assertIsNone(rows[0].get("ma_dvt_co_ban"))
            self.assertIsNone(rows[0].get("quy_doi"))


if __name__ == "__main__":
    unittest.main()
