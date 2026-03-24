# KiotViet → MTP

## Tổng quan

Script Python dùng để chuyển dữ liệu sản phẩm xuất từ KiotViet sang các file import của MTP.

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

Chạy script với file KiotViet:

```bash
python kiotviet_to_mtp.py --kiotviet /duong-dan/toi-file.xlsx
