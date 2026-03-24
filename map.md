### Các tiêu đề KiotViet đang hỗ trợ
- `Loại hàng`
- `Nhóm hàng(3 Cấp)` / `Nhóm hàng (3 Cấp)` / `Nhóm hàng`
- `Mã hàng`
- `Tên hàng`
- `Giá bán trước thuế` / `Giá bán`
- `Giá vốn`
- `Tồn kho`
- `ĐVT` / `Đơn vị tính`

### Từ KiotViet sang `MTP-NganhHang-LoaiHang.xls`
- `Loại hàng` → cột B `Tên ngành`
- Loại bỏ item trùng nhau
- Tự sinh `Mã ngành`

### Từ KiotViet sang `MTP-NhomHang.xls`
- `Nhóm hàng(3 Cấp)` → cột B `Tên nhóm`
- Loại bỏ item trùng nhau
- Cột C `Ngành hàng` luôn = `MD`
- Tự sinh `Mã nhóm`

### Từ KiotViet sang `MTP-SanPham.xls`
- `Nhóm hàng(3 Cấp)` → cột A `Nhóm hàng hóa`
- `Mã hàng` → cột B `Mã sản phẩm`
- `Tên hàng` → cột C `Tên sản phẩm`
- `ĐVT` → cột E `Đơn vị tính`
- `Giá vốn` → cột F `Giá nhập`
- `Giá bán trước thuế` → cột G `Giá bán`
- Cột H `Số lượng tồn tối thiểu` = `0`
- Cột I `Số lượng tồn tối đa` = `999999`

### Từ KiotViet sang `Mau-TonKhoDauKy.xls`
- `Mã hàng` → cột A `Mã sản phẩm`
- `Tồn kho` → cột B `Số lượng đầu kỳ`
- `Giá vốn` → cột C `Đơn giá đầu kỳ`