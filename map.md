### Các tiêu đề KiotViet đang hỗ trợ

#### Sản phẩm
- `Loại hàng`
- `Nhóm hàng(3 Cấp)` / `Nhóm hàng (3 Cấp)` / `Nhóm hàng`
- `Mã hàng`
- `Tên hàng`
- `Giá bán trước thuế` / `Giá bán`
- `Giá vốn`
- `Tồn kho`
- `ĐVT` / `Đơn vị tính`

#### Khách hàng
- `Mã khách hàng`
- `Tên khách hàng`
- `Điện thoại`
- `Địa chỉ`
- `Nợ cần thu hiện tại`

#### Nhà cung cấp
- `Mã nhà cung cấp`
- `Tên nhà cung cấp`
- `Email`
- `Điện thoại`
- `Địa chỉ`
- `Nợ cần trả hiện tại`

### Từ `DanhSachSanPham_*` sang `MTP_SP-NganhHang-LoaiHang.xls`
- `Loại hàng` → cột B `Tên ngành`
- Loại bỏ item trùng nhau
- Tự sinh `Mã ngành`

### Từ `DanhSachSanPham_*` sang `MTP_SP-NhomHang.xls`
- `Nhóm hàng(3 Cấp)` → cột B `Tên nhóm`
- Loại bỏ item trùng nhau
- Cột C `Ngành hàng` luôn = `MD`
- Tự sinh `Mã nhóm`

### Từ `DanhSachSanPham_*` sang `MTP_SP-SanPham.xls`
- `Nhóm hàng(3 Cấp)` → cột A `Nhóm hàng hóa`
- `Mã hàng` → cột B `Mã sản phẩm`
- `Tên hàng` → cột C `Tên sản phẩm`
- `ĐVT` → cột E `Đơn vị tính`
- `Giá vốn` → cột F `Giá nhập`
- `Giá bán trước thuế` / `Giá bán` → cột G `Giá bán`
- Cột H `Số lượng tồn tối thiểu` = `0`
- Cột I `Số lượng tồn tối đa` = `999999`

### Từ `DanhSachSanPham_*` sang `MTP_SP-TonKhoDauKy.xls`
- `Mã hàng` → cột A `Mã sản phẩm`
- `Tồn kho` → cột B `Số lượng đầu kỳ`
- `Giá vốn` → cột C `Đơn giá đầu kỳ`

### Từ `DanhSachKhachHang_*` sang `MTP_KH-NCC.xls`
- Cột C `Mã khách hàng` → cột A `Mã KH-NCC`
- Cột D `Tên khách hàng` → cột B `Tên KH-NCC`
- Cột F `Địa chỉ` → cột C `Địa chỉ`
- Cột E `Điện thoại` → cột H `Điện thoại`
- Cột E `Nhóm KH-NCC` mặc định = `KL`
- Cột F `Khách hàng(x)` mặc định = `TRUE`

### Từ `DanhSachNhaCungCap_*` sang `MTP_KH-NCC.xls`
- Cột A `Mã nhà cung cấp` → cột A `Mã KH-NCC`
- Cột B `Tên nhà cung cấp` → cột B `Tên KH-NCC`
- Cột E `Địa chỉ` → cột C `Địa chỉ`
- Cột D `Điện thoại` → cột H `Điện thoại`
- Cột C `Email` → cột I `Email`
- Cột E `Nhóm KH-NCC` mặc định = `NCC`
- Cột G `Nhà cung cấp(x)` mặc định = `TRUE`
- Nếu có nhiều file nguồn, dữ liệu được nối tiếp xuống dưới theo thứ tự truyền vào

### Từ `DanhSachKhachHang_*` sang `MTP_KH-CongNoDauKy.xls`
- Cột C `Mã khách hàng` → cột A `Mã KH - NCC(*)`
- Cột U `Nợ cần thu hiện tại` → cột C `Số tiền phải thu (*)`
- Cột D `Số tiền phải trả (*)` mặc định = `0`

### Từ `DanhSachNhaCungCap_*` sang `MTP_KH-CongNoDauKy.xls`
- Cột A `Mã nhà cung cấp` → cột A `Mã KH - NCC(*)`
- Cột I `Nợ cần trả hiện tại` → cột D `Số tiền phải trả (*)`
- Cột C `Số tiền phải thu (*)` mặc định = `0`
