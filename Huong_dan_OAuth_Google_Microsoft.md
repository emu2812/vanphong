# Hướng dẫn tự đăng ký app OAuth cho bản web “Văn phòng của Bùi”

> Lưu ý: file hướng dẫn Azure của nhánh Android (`~/workspace/office-app/Huong_dan_tao_app_Azure_Microsoft.md`)
> chưa có trong thư mục `office-app` tại thời điểm giao bản web, nên bản web tự viết bản
> hướng dẫn ngắn tương đương này đặt ngay trong `web/`. Khi nhánh Android giao file gốc,
> thay liên kết trong trang bằng file đó. Bản này không chứa Client ID, secret hay key nào —
> Bùi tự tạo, tự giữ.

## 1. Nguyên tắc

- Trang web không có sẵn Client ID nào và không bịa Client ID.
- Bùi tự đăng ký app, tự nhập Client ID vào mục **Cài đặt** trong trang.
- Luồng OAuth chỉ chạy trọn vẹn khi trang được phục vụ qua **https** và **redirect URI đã
  đăng ký khớp đúng địa chỉ trang hiện tại**. Mở file `index.html` bằng `file://` trên Safari
  thì nút đăng nhập sẽ báo rõ giới hạn này, không hiện lỗi khó hiểu.

## 2. Google (Drive)

1. Mở Google Cloud Console → tạo project mới → bật **Google Drive API**.
2. Cấu hình màn hình đồng ý OAuth (OAuth consent screen): chọn External, điền tên app.
3. Tạo **OAuth client ID** loại **Web application**.
4. Ở mục **Authorized JavaScript origins** và **Authorized redirect URIs**, thêm đúng địa chỉ
   https nơi Bùi đặt trang web này (ví dụ `https://<tên-miền-của-Bùi>/office-app/web/` và
   `https://<tên-miền-của-Bùi>/office-app/web/index.html`).
5. Copy **Client ID** (dạng `...apps.googleusercontent.com`) → dán vào ô **Google Client ID**
   trong mục Cài đặt của trang → bấm Lưu cài đặt → sang mục **Đăng nhập & Đồng bộ** bấm
   **Đăng nhập Google**. Quyền trang xin là `drive.file` (chỉ file do app này tạo/mở).

## 3. Microsoft (OneDrive, Azure)

1. Mở Azure Portal → **App registrations** → **New registration**.
2. Đặt tên app; mục **Redirect URI** chọn **Web**, nhập đúng địa chỉ https của trang này
   (giống mục 2, bước 4).
3. Ở **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** →
   thêm `User.Read` và `Files.ReadWrite` → **Grant admin consent** nếu được hỏi (tài khoản cá nhân
   thì tự đồng ý khi đăng nhập).
4. Copy **Application (client) ID** → dán vào ô **Microsoft Client ID** trong mục Cài đặt →
   Lưu cài đặt → sang mục **Đăng nhập & Đồng bộ** bấm **Đăng nhập Microsoft**.

## 4. OCR.space (tùy chọn, chỉ khi Tesseract.js không chạy)

1. Mở ocr.space → đăng ký tài khoản miễn phí → lấy API key miễn phí.
2. Dán key vào ô **OCR.space API key** trong mục Cài đặt. Trang không nhúng key nào sẵn.
3. Giới hạn bản miễn phí: giới hạn số lần gọi/ngày và kích thước file (xem bảng giá hiện tại
   trên ocr.space trước khi dùng nhiều).

## 5. Kiểm tra sau khi đăng nhập

- Đăng nhập xong ngay trong trang, trạng thái ở đầu trang đổi thành “đã đăng nhập”.
- Mở một tài liệu, bấm **Lên Drive** / **Lên OneDrive** trong danh sách file: chỉ khi API trả
  thành công, trạng thái mới đổi thành **Đã lên Drive / Đã lên OneDrive**; chưa đăng nhập
  hoặc lỗi mạng thì file ở **Chờ đồng bộ**, không báo thành công giả.
