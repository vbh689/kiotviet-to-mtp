"""Column mapping metadata shared by the GUI and Excel reader."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TypeAlias

ColumnMappings: TypeAlias = dict[str, dict[str, int]]


@dataclass(frozen=True)
class TargetUsage:
    template: str
    column: str
    label: str


@dataclass(frozen=True)
class FieldMapping:
    field: str
    label: str
    targets: tuple[TargetUsage, ...]


SOURCE_TYPE_LABELS = {
    "product": "Sản phẩm",
    "customer": "Khách hàng",
    "provider": "Nhà cung cấp",
}

MAPPING_METADATA: dict[str, tuple[FieldMapping, ...]] = {
    "product": (
        FieldMapping(
            "loai_hang",
            "Loại hàng",
            (TargetUsage("MTP_SP-NganhHang-LoaiHang.xls", "B", "Tên ngành"),),
        ),
        FieldMapping(
            "nhom_hang",
            "Nhóm hàng",
            (
                TargetUsage("MTP_SP-NhomHang.xls", "B", "Tên nhóm"),
                TargetUsage("MTP_SP-SanPham.xls", "A", "Nhóm hàng hóa"),
            ),
        ),
        FieldMapping(
            "ma_hang",
            "Mã hàng",
            (
                TargetUsage("MTP_SP-SanPham.xls", "B", "Mã sản phẩm"),
                TargetUsage("MTP_SP-TonKhoDauKy.xls", "A", "Mã sản phẩm"),
            ),
        ),
        FieldMapping(
            "ten_hang",
            "Tên hàng",
            (TargetUsage("MTP_SP-SanPham.xls", "C", "Tên sản phẩm"),),
        ),
        FieldMapping(
            "don_vi_tinh",
            "ĐVT",
            (TargetUsage("MTP_SP-SanPham.xls", "E", "Đơn vị tính"),),
        ),
        FieldMapping(
            "gia_von",
            "Giá vốn",
            (
                TargetUsage("MTP_SP-SanPham.xls", "F", "Giá nhập"),
                TargetUsage("MTP_SP-TonKhoDauKy.xls", "C", "Đơn giá đầu kỳ"),
            ),
        ),
        FieldMapping(
            "gia_ban",
            "Giá bán",
            (TargetUsage("MTP_SP-SanPham.xls", "G", "Giá bán"),),
        ),
        FieldMapping(
            "ton_kho",
            "Tồn kho",
            (TargetUsage("MTP_SP-TonKhoDauKy.xls", "B", "Số lượng đầu kỳ"),),
        ),
    ),
    "customer": (
        FieldMapping(
            "ma_khach_hang",
            "Mã khách hàng",
            (
                TargetUsage("MTP_KH-NCC.xls", "A", "Mã KH-NCC"),
                TargetUsage("MTP_KH-CongNoDauKy.xls", "A", "Mã KH - NCC(*)"),
            ),
        ),
        FieldMapping(
            "ten_khach_hang",
            "Tên khách hàng",
            (TargetUsage("MTP_KH-NCC.xls", "B", "Tên KH-NCC"),),
        ),
        FieldMapping(
            "dia_chi",
            "Địa chỉ",
            (TargetUsage("MTP_KH-NCC.xls", "C", "Địa chỉ"),),
        ),
        FieldMapping(
            "dien_thoai",
            "Điện thoại",
            (TargetUsage("MTP_KH-NCC.xls", "H", "Điện thoại"),),
        ),
        FieldMapping(
            "no_can_thu_hien_tai",
            "Nợ cần thu hiện tại",
            (TargetUsage("MTP_KH-CongNoDauKy.xls", "C", "Số tiền phải thu (*)"),),
        ),
    ),
    "provider": (
        FieldMapping(
            "ma_nha_cung_cap",
            "Mã nhà cung cấp",
            (
                TargetUsage("MTP_KH-NCC.xls", "A", "Mã KH-NCC"),
                TargetUsage("MTP_KH-CongNoDauKy.xls", "A", "Mã KH - NCC(*)"),
            ),
        ),
        FieldMapping(
            "ten_nha_cung_cap",
            "Tên nhà cung cấp",
            (TargetUsage("MTP_KH-NCC.xls", "B", "Tên KH-NCC"),),
        ),
        FieldMapping(
            "dia_chi",
            "Địa chỉ",
            (TargetUsage("MTP_KH-NCC.xls", "C", "Địa chỉ"),),
        ),
        FieldMapping(
            "dien_thoai",
            "Điện thoại",
            (TargetUsage("MTP_KH-NCC.xls", "H", "Điện thoại"),),
        ),
        FieldMapping(
            "email",
            "Email",
            (TargetUsage("MTP_KH-NCC.xls", "I", "Email"),),
        ),
        FieldMapping(
            "no_can_tra_hien_tai",
            "Nợ cần trả hiện tại",
            (TargetUsage("MTP_KH-CongNoDauKy.xls", "D", "Số tiền phải trả (*)"),),
        ),
    ),
}


def required_fields_for(source_type: str) -> tuple[str, ...]:
    return tuple(item.field for item in MAPPING_METADATA[source_type])
