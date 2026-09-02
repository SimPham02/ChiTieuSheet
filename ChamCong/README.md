# Ứng Dụng Chấm Công & Nhật Ký Đi Làm (Google Sheets Web App)

Ứng dụng ghi nhận chấm công cá nhân kết nối trực tiếp với Google Sheets, hiển thị dưới dạng **cuốn lịch trực quan (Interactive Calendar View)** với giao diện hiện đại, trực quan, hỗ trợ theo dõi ngày đi làm, ngày nghỉ, tăng ca (OT), ghi chú công việc và thống kê chuyên cần hàng tháng.

---

## ⏰ Quy Chuẩn Giờ Làm Việc & OT Mặc Định

- **Giờ làm việc tiêu chuẩn**: `08:00` sáng đến `17:00` chiều (Tính `1.0` công chuẩn).
- **Tăng ca (OT) mặc định**: Mặc định đến `20:00` tối (Tính `3.0` giờ OT).
- **Nửa ngày**: `08:00` đến `12:00` (Tính `0.5` công).
- Khi bấm chọn trạng thái hoặc các nút preset trên giao diện, hệ thống tự động điền sẵn giờ vào, giờ ra và số giờ OT tương ứng.

---

## 🌟 Tính Năng Nổi Bật

1. **Giao diện Cuốn Lịch (Calendar View)**:
   - Hiển thị đầy đủ các ngày trong tháng theo tuần (Thứ 2 đến Chủ nhật).
   - Đánh dấu nổi bật ngày "Hôm nay", phân biệt cuối tuần (T7, CN).
   - Hiển thị trực quan Badge trạng thái:
     - 🟢 **Đi làm** (1.0 công / 0.5 công)
     - 🟠 **Tăng ca (OT)** (Ví dụ: +3h đến 20:00)
     - 🔴 **Nghỉ phép** / ⚪ **Nghỉ không lương** / 🔵 **Nghỉ lễ**
   - Hiển thị dòng ghi chú tóm tắt (Note) ngay trên từng ô ngày của lịch.

2. **Chấm Công & Note Nhanh (1 Chạm)**:
   - Nhấp vào bất kỳ ngày nào trên lịch hoặc bấm nút **"Chấm công hôm nay"**.
   - Chọn nhanh trạng thái (Đi làm, Nửa ngày, Làm + OT, Nghỉ phép...).
   - Nhập nhanh số giờ OT với các nút bấm tiện lợi: `+1h (18h)`, `+2h (19h)`, `+3h (20h Mặc định)`, `+4h (21h)`.
   - Nhập giờ vào / ra và dòng ghi chú công việc trong ngày.

3. **Chấm Nhanh Toàn Tháng (Batch Fill T2 - T6)**:
   - 1 nút bấm tự động điền trạng thái "Đi làm (08:00 - 17:00, 1 công)" cho tất cả các ngày trong tuần (T2 đến T6) của tháng hiện tại chưa được chấm.

4. **Thống Kê KPI & Đánh Giá Hàng Tháng**:
   - **Tổng số công**: Theo dõi tổng công thực tế và tiến độ so với công chuẩn (22 công).
   - **Tăng ca (OT)**: Tổng số giờ OT và số ngày có tăng ca trong tháng.
   - **Ngày nghỉ**: Tổng hợp số ngày nghỉ phép và nghỉ không lương.
   - **Tỉ lệ chuyên cần (%)**: Đánh giá tự động mức độ hoàn thành công việc.

5. **Linh Hoạt Chế Độ Xem & Xuất Dữ Liệu**:
   - Chuyển đổi giữa chế độ **Lịch tháng** và **Nhật ký dạng bảng (Table Log)**.
   - Xuất dữ liệu chấm công ra file **CSV** để báo cáo hoặc tính lương.

6. **Chế Độ Chạy Mượt Mà**:
   - Chạy trực tiếp trên Google Apps Script (lưu vào Google Sheet).
   - Tự động chuyển sang chế độ **Local Demo (Offline)** khi mở file `index.html` trực tiếp trên trình duyệt mà không cần cài đặt thêm.

---

## 🚀 Hướng Dẫn Cài Đặt & Triển Khai Lên Google Sheets

### Bước 1: Tạo Google Sheet mới
1. Truy cập [Google Sheets](https://sheets.new) và tạo một bảng tính mới.
2. Đổi tên bảng tính (ví dụ: `Chấm Công & Đi Làm 2026`).

### Bước 2: Mở Trình chỉnh sửa Apps Script
1. Trên thanh menu của Google Sheet, chọn **Tiện ích mở rộng** (Extensions) > **Apps Script**.
2. Đổi tên dự án Apps Script thành `ChamCongApp`.

### Bước 3: Thêm mã nguồn
1. Tại file `Mã.gs` (hoặc `Code.gs`):
   - Xóa toàn bộ mã mặc định.
   - Sao chép toàn bộ nội dung trong file [`ChamCong/code.gs`](./code.gs) và dán vào.
2. Thêm file HTML:
   - Nhấn vào dấu `+` cạnh mục **Tệp** (Files) > Chọn **HTML**.
   - Đặt tên file là `index` (Google Apps Script sẽ tự động tạo file `index.html`).
   - Xóa mã mặc định, sao chép toàn bộ nội dung file [`ChamCong/index.html`](./index.html) và dán vào.
3. Nhấn biểu tượng **Lưu** (Save / Ctrl + S).

### Bước 4: Triển khai Web App (Deploy)
1. Ở góc trên bên phải, nhấn nút **Triển khai** (Deploy) > Chọn **Triển khai mới** (New deployment).
2. Nhấn vào biểu tượng bánh răng ⚙️ cạnh **Chọn loại** (Select type) > Chọn **Ứng dụng web** (Web app).
3. Điền thông tin:
   - **Mô tả** (Description): `Phiên bản 1.0`
   - **Thực thi dưới dạng** (Execute as): `Tôi` (Me / tài khoản Google của bạn)
   - **Người có quyền truy cập** (Who has access):
     - Chọn `Bất kỳ ai có tài khoản Google` (Anyone with Google account) hoặc `Chỉ mình tôi` (Only myself) hoặc `Bất kỳ ai` (Anyone).
4. Nhấn **Triển khai** (Deploy).
5. Cấp quyền truy cập khi được yêu cầu (Google sẽ yêu cầu xác thực quyền đọc/ghi vào Google Sheets của bạn).
6. Sao chép đường dẫn **URL ứng dụng web** (Web app URL) và mở trên trình duyệt hoặc bookmark trên điện thoại/máy tính để sử dụng hàng ngày!

---

## 📊 Cấu Trúc Dữ Liệu Sheet `CHAM_CONG`

Bảng tính sẽ tự động được khởi tạo các cột với định dạng chuẩn:

| Cột | Tên cột | Ý nghĩa | Mặc định / Ví dụ |
|---|---|---|---|
| A | **ID** | Mã định danh | `cc_2026-09-02` |
| B | **Ngày** | Ngày làm việc (YYYY-MM-DD) | `2026-09-02` |
| C | **Thứ** | Thứ trong tuần | `Thứ Tư` |
| D | **Trạng thái** | Trạng thái công việc | `Đi làm`, `OT`, `Nửa ngày`, `Nghỉ phép` |
| E | **Số công** | Điểm công quy đổi | `1.0`, `0.5`, `0` |
| F | **Giờ OT** | Số giờ làm thêm | `3.0` (đến 20h) |
| G | **Giờ vào** | Giờ check-in | `08:00` |
| H | **Giờ ra** | Giờ check-out | `17:00` (hoặc `20:00` nếu OT) |
| I | **Ghi chú** | Chi tiết công việc / Ghi chú ngày | `Tăng ca hoàn thành sprint` |
| J | **Thời gian cập nhật** | Thời điểm ghi nhận | `2026-09-02 17:45:00` |
