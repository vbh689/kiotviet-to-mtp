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
pip install xlrd xlwt openpyxl PyQt6 pyinstaller
```

## Sử dụng

Chạy bằng giao diện cửa sổ kéo-thả (GUI) mặc định:

```bash
python kiotviet_to_mtp.py
```

Trong GUI:

1. Kéo thả một hoặc nhiều file `.xlsx` xuất từ KiotViet.
2. Kiểm tra hộp thoại chọn cột. Ứng dụng sẽ tự chọn các cột khớp với tiêu đề mặc định hiện có.
3. Đổi cột nguồn nếu file KiotViet của bạn đổi tên tiêu đề hoặc sắp xếp lại cột.
4. Bấm OK để chuyển đổi. Mọi mapping tùy chỉnh chỉ áp dụng cho lần chạy hiện tại và không được lưu.

Hoặc chạy dòng lệnh (CLI) tự động với một hoặc nhiều file KiotViet:

```bash
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachSanPham.xlsx
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachKhachHang.xlsx /duong-dan/DanhSachNhaCungCap.xlsx
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachSanPham.xlsx /duong-dan/DanhSachKhachHang.xlsx /duong-dan/DanhSachNhaCungCap.xlsx
```

Dùng `--outdir` để đổi thư mục xuất:

```bash
python kiotviet_to_mtp.py --kiotviet /duong-dan/DanhSachKhachHang.xlsx /duong-dan/DanhSachNhaCungCap.xlsx --outdir ./output
```

CLI vẫn dùng cơ chế tự nhận diện cột bằng tiêu đề mặc định/alias trong `map.md` và không hỏi mapping tùy chỉnh.

## File đầu ra

Tùy theo loại file nguồn được truyền vào, script sẽ sinh các file tương ứng:

- `MTP_SP-NganhHang-LoaiHang.xls`
- `MTP_SP-NhomHang.xls`
- `MTP_SP-SanPham.xls`
- `MTP_SP-TonKhoDauKy.xls`
- `MTP_KH-NCC.xls`
- `MTP_KH-CongNoDauKy.xls`

## Đóng gói ứng dụng (Compile Guide)

Bạn có thể đóng gói toàn bộ chương trình thành 1 file chạy duy nhất (Executable) bao gồm sẵn giao diện (GUI) và thư mục `templates/`. Tùy theo hệ điều hành hiện tại, vui lòng chạy các lệnh sau:

### Cho MacOS / Linux:
```bash
pyinstaller --onedir --windowed --add-data "templates:templates" kiotviet_to_mtp.py
```
> Lưu ý: dấu ngăn cách trong `--add-data` của MacOS là dấu hai chấm `:`. File đầu ra nằm ở `dist/kiotviet_to_mtp`.

### Cho Windows:
```bash
pyinstaller --onefile --windowed --add-data "templates;templates" kiotviet_to_mtp.py
```
> Lưu ý: dấu ngăn cách trong `--add-data` của Windows là dấu chấm phẩy `;`. File đầu ra nằm ở `dist\kiotviet_to_mtp.exe`.
