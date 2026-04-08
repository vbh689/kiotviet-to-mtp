# KiotViet → MTP

## Tổng quan

Script Python dùng để chuyển dữ liệu xuất từ KiotViet sang các file import của MTP:

- Sản phẩm
- Khách hàng
- Nhà cung cấp
- Công nợ đầu kỳ KH/NCC

## Cài đặt

Tạo và kích hoạt môi trường ảo:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Cài các thư viện cần thiết:

```bash
pip install xlrd xlwt openpyxl
```

## Sử dụng

Chạy script với một hoặc nhiều file KiotViet:

```bash
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachSanPham.xlsx
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachKhachHang.xlsx /duong-dan/DanhSachNhaCungCap.xlsx
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachSanPham.xlsx /duong-dan/DanhSachKhachHang.xlsx /duong-dan/DanhSachNhaCungCap.xlsx
```

Dùng `--outdir` để đổi thư mục xuất:

```bash
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachKhachHang.xlsx /duong-dan/DanhSachNhaCungCap.xlsx --outdir ./output
```

## File đầu ra

Tùy theo loại file nguồn được truyền vào, script sẽ sinh các file tương ứng:

- `MTP_SP-NganhHang-LoaiHang.xls`
- `MTP_SP-NhomHang.xls`
- `MTP_SP-SanPham.xls`
- `MTP_SP-TonKhoDauKy.xls`
- `MTP_KH-NCC.xls`
- `MTP_KH-CongNoDauKy.xls`
