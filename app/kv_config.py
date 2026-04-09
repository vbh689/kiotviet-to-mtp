"""Static conversion configuration and supported header names."""

from __future__ import annotations

from .kv_utils import flatten_aliases


# Supported column aliases for each KiotViet export type.
# Header matching later normalizes text, so these lists can include small naming variants.
PRODUCT_HEADER_ALIASES = {
    "loai_hang": ["Loại hàng"],
    "nhom_hang": ["Nhóm hàng(3 Cấp)", "Nhóm hàng (3 Cấp)", "Nhóm hàng"],
    "ma_hang": ["Mã hàng"],
    "ten_hang": ["Tên hàng"],
    "gia_ban": ["Giá bán trước thuế", "Giá bán"],
    "gia_von": ["Giá vốn"],
    "ton_kho": ["Tồn kho"],
    "don_vi_tinh": ["ĐVT", "Đơn vị tính"],
}

CUSTOMER_HEADER_ALIASES = {
    "ma_khach_hang": ["Mã khách hàng"],
    "ten_khach_hang": ["Tên khách hàng"],
    "dien_thoai": ["Điện thoại"],
    "dia_chi": ["Địa chỉ"],
    "no_can_thu_hien_tai": ["Nợ cần thu hiện tại"],
}

PROVIDER_HEADER_ALIASES = {
    "ma_nha_cung_cap": ["Mã nhà cung cấp"],
    "ten_nha_cung_cap": ["Tên nhà cung cấp"],
    "email": ["Email"],
    "dien_thoai": ["Điện thoại"],
    "dia_chi": ["Địa chỉ"],
    "no_can_tra_hien_tai": ["Nợ cần trả hiện tại"],
}

# Filename prefixes are checked first when guessing whether an input sheet is
# product, customer, or provider data.
SOURCE_PREFIXES = {
    "product": "DanhSachSanPham",
    "customer": "DanhSachKhachHang",
    "provider": "DanhSachNhaCungCap",
}

# Some templates have old and new filenames, so each logical output can point to
# more than one accepted candidate.
TEMPLATE_CANDIDATES = {
    "product_nganh": ["MTP_SP-NganhHang-LoaiHang.xls", "MTP-NganhHang-LoaiHang.xls"],
    "product_nhom": ["MTP_SP-NhomHang.xls", "MTP-NhomHang.xls"],
    "product_sanpham": ["MTP_SP-SanPham.xls", "MTP-SanPham.xls"],
    "product_tonkho": ["MTP_SP-TonKhoDauKy.xls", "Mau-TonKhoDauKy.xls"],
    "kh_ncc": ["MTP_KH-NCC.xls"],
    "kh_congno": ["MTP_KH-CongNoDauKy.xls"],
}

KNOWN_PRODUCT_HEADERS = flatten_aliases(PRODUCT_HEADER_ALIASES)
KNOWN_CUSTOMER_HEADERS = flatten_aliases(CUSTOMER_HEADER_ALIASES)
KNOWN_PROVIDER_HEADERS = flatten_aliases(PROVIDER_HEADER_ALIASES)
