# Văn phòng của Bùi — Bản web (PWA) 1.1

## Vì sao là bản web, không phải app App Store thật

**Không hứa hẹn app App Store.** Một app iOS thật phát hành trên App Store cần **tài khoản
Apple Developer trả phí + máy Mac để ký và đóng gói**, mà môi trường làm bản này không có
cả hai. Vì vậy đầu ra cho iPhone/iPad là **bản web (PWA)**: mở `index.html` bằng Safari,
và khi được phục vụ qua https thì cài được như app qua **Chia sẻ → Thêm vào Màn hình chính**
(kèm `manifest.json` + service worker `sw.js`).

## Chạy như thế nào

- **Mở trực tiếp:** mở file `index.html` trên Safari iPhone/iPad là dùng được soạn thảo,
  bảng tính, lưu/mở lại, xuất file, chụp ảnh tạo PDF (phần không cần CDN).
- **Cài như PWA / chạy đầy đủ:** đưa cả thư mục `web/` lên một host **https** (ví dụ GitHub
  Pages), mở link https đó, rồi Thêm vào Màn hình chính. Lúc này service worker mới đăng ký
  được — **service worker không chạy ở chế độ `file://`**.
- **HTML/JS thuần, không framework build**, không cần cài đặt, không thu thập dữ liệu,
  không nhúng key/secret, không ghi sẵn nội dung mẫu.

## Tính năng bản 1.0 — đạt / chưa đạt

| Tính năng | Trạng thái |
|---|---|
| 1. Soạn văn bản (contenteditable + font, cỡ, đậm/nghiêng/gạch chân/gạch ngang, màu chữ, tô sáng, khoảng cách dòng, căn lề, gạch đầu dòng, tiêu đề; lưu localStorage, mở lại đúng định dạng; xuất .txt và .docx tối giản tự viết zip stored + CRC32, giữ font/cỡ/đậm/nghiêng/gạch chân/gạch ngang/màu/căn lề) | ✅ Đạt (mã đã viết; cần Bùi mở trên Safari thật để bấm thử) |
| 2. Bảng tính (lưới nhập số/chữ, `=SUM(A1:A3)`, `=AVERAGE(A1:A3)`, tham chiếu ô/vùng, tính lại khi sửa; lưu/mở lại; xuất .csv và .xlsx tối giản tự viết) | ✅ Đạt (.xlsx làm được nên có luôn, không chỉ .csv) |
| 2a. Trình chiếu — MỚI bản 1.1 (nhiều slide, mỗi slide tiêu đề + gạch đầu dòng, chèn ảnh nén tối đa 1280px; trình chiếu toàn màn hình, vuốt trái/phải hoặc nút lớn để chuyển slide; lưu/mở lại localStorage như tài liệu/bảng tính; xuất .pptx tối giản tự viết zip + xml, JS thuần, không thư viện ngoài) | ✅ Đạt (mã đã viết; .pptx đã kiểm chứng mở được bằng LibreOffice Impress và trích đúng chữ tiếng Việt — xem dưới; vẫn cần Bùi bấm thử vuốt/nút trên Safari thật) |
| 3. PDF (chọn file, xem từng trang bằng pdf.js qua CDN; tô sáng, ghi chú, vẽ tay; chú thích lưu localStorage kèm tên file, mở lại thấy đủ) | ✅ Đạt, với giới hạn dưới đây |
| 3a. OCR (trang PDF thành ảnh hoặc ảnh riêng → văn bản sửa được, vie+eng) | ✅ Đạt bằng Tesseract.js (vie+eng) làm chính, OCR.space làm dự phòng — xem mục OCR |
| 3b. Chụp ảnh tạo PDF thật (camera/getUserMedia hoặc input capture, xem lại, lên/xuống/xóa trang, tạo PDF bằng JS thuần nhúng JPEG, lưu vào danh sách file để đồng bộ, mở lại bằng OCR) | ✅ Đạt (không cần thư viện ngoài) |
| 4. Dịch Việt ↔ Anh/Trung/Hàn/Nhật qua MyMemory (fetch, không key); dịch đoạn đã chọn rồi chèn ngay dưới đoạn gốc; mất mạng báo rõ, không bịa | ✅ Đạt |
| 5. Quản lý file (danh sách localStorage; nhập txt, docx văn bản thuần, csv, pdf; xuất/tải về; chia sẻ qua Web Share API khi trình duyệt hỗ trợ) | ✅ Đạt, với giới hạn docx dưới đây |
| 6. Font tiếng Việt (chỉ 5 font đã kiểm chứng) | ✅ Đạt — xem mục Font |
| 7. Đăng nhập Google & Microsoft **ngay trong trang này** (mục “Đăng nhập & Đồng bộ”, 2 nút; đăng nhập xong lưu lên Drive/OneDrive và tải về mở lại ngay trong trang; trạng thái hiện trong danh sách file) | ✅ Đạt về mặt mã (luồng OAuth chỉ chạy trọn vẹn qua https + redirect URI đã đăng ký — xem mục OAuth) |

## Giới hạn phải đọc kỹ

- **PDF bản 1.0 CHƯA ghi đè vào ruột file PDF gốc.** Chú thích chỉ lưu kèm tên file trong
  trình duyệt. Mở lại đúng file đó thì thấy đủ chú thích; file gốc giữ nguyên.
- **Nhập docx chỉ đọc văn bản thuần** (không giữ nguyên định dạng phức tạp, bảng, ảnh).
  Giải nén bằng JS thuần: entry stored lấy thẳng, entry deflate dùng `DecompressionStream`
  nếu trình duyệt có; trình duyệt không có thì báo rõ giới hạn thay vì im lặng.
- **File PDF gốc không lưu vào localStorage** (file có thể lớn). Chú thích lưu theo tên file;
  muốn xem lại chú thích thì chọn lại đúng file PDF đó. Riêng **PDF tạo bằng camera được lưu
  Blob vào IndexedDB** nên mở lại được từ danh sách file mà không cần chọn lại.
- **Đăng xuất Microsoft:** nút đăng xuất xóa token trong trang/phiên; muốn thoát hẳn tài khoản
  Microsoft trên trình duyệt thì đăng xuất thêm ở tài khoản Microsoft nếu trình duyệt còn nhớ phiên.

## Font chữ — danh sách cuối cùng và cách kiểm chứng

Danh sách cuối cùng, **chỉ 5 font**, kèm font hệ thống dự phòng
(`-apple-system, "Segoe UI", sans-serif`):

1. Be Vietnam Pro — 2. Inter — 3. Roboto — 4. Noto Sans — 5. Noto Serif

Mọi font khác không chắc hỗ trợ tiếng Việt đều đã loại.

Cách kiểm chứng trước khi giao:

1. **Đối chiếu metadata Google Fonts:** cả 5 family đều có subset `vietnamese` (Google Fonts
   trả `unicode-range` phủ dấu tiếng Việt cho các family này; Be Vietnam Pro vốn thiết kế
   cho tiếng Việt).
2. **Đoạn kiểm chứng trong trang** (mục Cài đặt): “Tiếng Việt có dấu: à á ả ã ạ, ồ ộ ổ,
   ứ ự ử, đ, ê, ơ, ư — Bác sĩ Nguyễn Văn Đô đi học ở Hà Nội” — chứa đủ thanh dấu, nguyên âm
   đôi và chữ đ — kèm nút đổi qua **từng font** trong danh sách để nhìn bằng mắt.
3. **Đổi font áp dụng ngay** lên đoạn đang chọn và giữ đúng khi lưu/mở lại (HTML lưu trong
   localStorage) và khi xuất .docx (ghi đúng tên font vào `w:rFonts`).

Lưu ý trung thực: môi trường làm bản web không có Safari iPhone thật, nên bước **nhìn bằng
mắt cuối cùng cần Bùi kiểm tra trên máy thật** bằng các nút đổi font trong trang. Về mặt
subset thì cả 5 font đều phủ đầy đủ dấu tiếng Việt nên không font nào vỡ dấu do thiếu glyph.

## OCR — đang chạy bằng cách nào

- **Chính: Tesseract.js qua CDN, chạy ngay trong trình duyệt, nhận dạng `vie+eng`**
  (tiếng Việt có dấu + tiếng Anh). Cần mạng lần đầu để tải script và dữ liệu vie+eng;
  lần sau trình duyệt có thể dùng bản đã lưu trong bộ nhớ đệm. Kết quả vào ô văn bản
  **sửa trực tiếp được**, rồi sao chép, chèn vào tài liệu, hoặc lưu thành tài liệu mới.
  Nhận dạng sai thì Bùi sửa tay, không khóa cứng.
- **Dự phòng: OCR.space qua mạng (language `vie`)** — chỉ dùng khi Tesseract.js không chạy
  được trên Safari. Cần Bùi tự tạo key miễn phí ở ocr.space rồi nhập vào Cài đặt; trang
  **không nhúng key trả phí nào**. Giới hạn bản miễn phí: giới hạn số lần gọi/ngày và kích
  thước file (xem bảng hiện tại trên ocr.space).
- **Không bịa kết quả OCR:** lỗi mạng / lỗi dịch vụ thì ô kết quả để trống kèm thông báo rõ.

## Phần nào cần mạng, phần nào chạy ngoại tuyến

- **Ngoại tuyến (sau khi trang đã tải):** soạn thảo, bảng tính, trình chiếu, lưu/mở lại, xuất
  txt/docx/csv/xlsx/pptx, chú thích PDF (khi pdf.js đã tải), chụp ảnh tạo PDF.
- **Cần mạng:** tải pdf.js lần đầu, tải Tesseract.js + dữ liệu vie+eng lần đầu, tải Google
  Fonts lần đầu (sau đó trình duyệt lưu đệm), dịch MyMemory, đăng nhập và đồng bộ
  Drive/OneDrive, tải thư viện OAuth (Google Identity Services, MSAL.js) khi đăng nhập.

## OAuth — đăng nhập ngay trong trang, không app khác

- Trong trang có sẵn mục **“Đăng nhập & Đồng bộ”** với 2 nút **Đăng nhập Google** và
  **Đăng nhập Microsoft**. Đăng nhập xong, Bùi lưu file lên Drive/OneDrive và tải file về
  mở lại **ngay trong trang đó**; trạng thái đồng bộ hiện ngay trong danh sách file ở mục
  **Quản lý file**.
- **KHÔNG bịa Client ID:** trang không có sẵn Client ID nào. Bùi tự đăng ký app rồi nhập
  Client ID vào mục Cài đặt (xem `Huong_dan_OAuth_Google_Microsoft.md` trong cùng thư mục;
  file hướng dẫn Azure của nhánh Android chưa có trong thư mục `office-app` tại thời điểm
  giao bản web, nên bản web tự viết bản hướng dẫn ngắn tương đương này — khi nhánh Android
  giao file gốc thì thay liên kết bằng file đó).
- Khi chưa nhập Client ID, nút đăng nhập báo rõ cần làm bước đó trước. Khi đã nhập, luồng
  OAuth chạy bằng thư viện qua CDN (Google Identity Services, MSAL.js) với redirect về đúng
  trang hiện tại.
- **Giới hạn OAuth:** luồng chỉ chạy trọn vẹn khi trang được phục vụ qua **https** với
  redirect URI đã đăng ký. Mở `file://` trên Safari thì nút đăng nhập báo rõ giới hạn này
  chứ không hiện lỗi khó hiểu.
- Nguyên tắc như bản Android: file luôn lưu ở máy (localStorage/IndexedDB) trước, có tài
  khoản và có mạng mới đồng bộ. Trạng thái từng file: **Đã lưu máy / Chờ đồng bộ /
  Đã lên Drive / Đã lên OneDrive** — trạng thái Drive/OneDrive chỉ chuyển thành “đã lên”
  khi đồng bộ qua API thật sự thành công.

## Dùng bằng cảm ứng (điện thoại, không chuột) — kết quả rà soát

Bản web thiên về dùng trên điện thoại, **mọi thao tác làm được bằng ngón tay**:

- **Không tính năng nào bắt buộc chuột hay bàn phím:** không menu chỉ hiện khi rê chuột
  (không dùng `:hover` để hiện nút quan trọng — trong CSS không còn quy tắc hover nào hiện/
  ẩn nút), không bấm chuột phải, không thao tác cần con trỏ chính xác từng pixel. Bàn phím
  chỉ dùng để gõ chữ; mọi nút chức năng đều bấm được bằng tay.
- **Vùng chạm tối thiểu 48px:** tất cả nút, ô nhập, ô chọn đều `min-height: 48px`; nút trong
  bottom sheet là 56px.
- **Chọn font/cỡ/màu/tiêu đề/mở tài liệu bằng bảng chọn trượt từ dưới lên (bottom sheet)**
  thay vì menu thả xuống nhỏ; soạn tài liệu chạm vào là gõ.
- **Bảng tính:** chạm một lần để chọn ô, chạm hai lần liên tiếp để sửa ô, vuốt để cuộn.
- **PDF:** chụm hai ngón để phóng to/thu nhỏ, vẽ chú thích bằng ngón tay, chạm để thêm ghi chú.
- **Các luồng đã rà lại từng luồng, khẳng định làm được chỉ bằng chạm, không cần chuột:**
  soạn tài liệu (chạm → gõ → bấm Lưu/Xuất) ✅; bảng tính (chạm chọn → chạm đúp sửa → Lưu/Xuất) ✅;
  PDF (chọn file → vuốt/chụm → chạm/vẽ chú thích) ✅; OCR (bấm OCR trang/ảnh → sửa kết quả
  bằng bàn phím cảm ứng → chèn/lưu) ✅; chụp ảnh tạo PDF (mở camera hoặc chọn ảnh → bấm
  Lên/Xuống/Xóa → bấm Tạo PDF → bấm Mở để OCR) ✅; dịch (gõ → bấm Dịch → bấm chèn) ✅;
  đăng nhập (nhập Client ID → bấm Đăng nhập → bấm Lên Drive/OneDrive trong danh sách file) ✅;
  đồng bộ (trạng thái hiện ngay trong danh sách file, không cần thao tác chuột) ✅.
- Lưu ý trung thực: rà soát này là rà logic mã + cấu trúc giao diện (nút, kích thước chạm,
  sự kiện chạm/chụm) trong môi trường không có iPhone thật; Bùi vẫn nên bấm thử một vòng
  trên iPhone thật để xác nhận cảm giác chạm cuối cùng (đặc biệt chụm phóng PDF và camera).

## Kiểm chứng bản 1.1 (trình chiếu + .pptx)

- `node --check app.js` đạt.
- Chạy app.js trong harness node với DOM giả: tạo slide → sửa tiêu đề/nội dung → thêm slide → đẩy slide lên → xóa slide → lưu → mở lại (tiêu đề/nội dung khôi phục đúng) → trình chiếu (hiện đúng slide, nút Trước/Sau kẹt đúng ở hai đầu, Thoát ẩn màn hình) — tất cả pass.
- File `.pptx` do chính hàm `buildPptxBlob` trong app.js tạo (2 slide, 1 slide có ảnh JPEG): kiểm tra zip + mọi file xml đều well-formed; **LibreOffice Impress mở được và chuyển sang PDF 2 trang, trích đúng chữ** “Tiêu đề có dấu: à á ả ã ạ, ồ ộ ổ, đ” và “Slide có ảnh • Mô tả ảnh”. PowerPoint/WPS thật chưa kiểm chứng trực tiếp (không có máy đó ở đây) — Bùi mở thử file .pptx xuất từ trang trên máy thật để chốt cuối.
- Giới hạn trình chiếu: bài trình bày chỉ lưu trong trình duyệt (localStorage); ảnh nén tối đa 1280px nên ảnh gốc rất lớn sẽ nhỏ lại; .pptx tối giản (tiêu đề + gạch đầu dòng + ảnh, không có hiệu ứng chuyển slide, không có mẫu nền cầu kỳ, không nhập ngược .pptx về sửa).

## File trong thư mục web/

- `index.html` — toàn bộ giao diện (8 mục, tiếng Việt; thêm mục Trình chiếu + màn hình trình chiếu toàn màn hình).
- `style.css` — giao diện, quy tắc cảm ứng 48px, bottom sheet, thẻ slide và màn hình trình chiếu.
- `app.js` — toàn bộ logic (soạn thảo, bảng tính, trình chiếu + xuất .pptx, PDF, OCR, camera tạo PDF, dịch, file, OAuth).
- `manifest.json`, `sw.js`, `icon.svg` — PWA (chỉ có tác dụng đầy đủ qua https).
- `Huong_dan_OAuth_Google_Microsoft.md` — hướng dẫn tự đăng ký app OAuth (bản web tự viết).
- `README.md` — file này.
