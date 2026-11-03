# Văn phòng của Bùi — Bản web (PWA) 1.2

## Vì sao là bản web, không phải app App Store thật

**Không hứa hẹn app App Store.** Một app iOS thật phát hành trên App Store cần **tài khoản
Apple Developer trả phí + máy Mac để ký và đóng gói**, mà môi trường làm bản này không có
cả hai. Vì vậy đầu ra cho iPhone/iPad là **bản web (PWA)**: mở `index.html` bằng Safari,
và khi được phục vụ qua https thì cài được như app qua **Chia sẻ → Thêm vào Màn hình chính**
(kèm `manifest.json` + service worker `sw.js`, cache `van-phong-bui-v3`).

## Chạy như thế nào

- **Mở trực tiếp:** mở file `index.html` trên Safari iPhone/iPad là dùng được soạn thảo,
  bảng tính, lưu/mở lại, xuất file, chụp ảnh tạo PDF (phần không cần CDN).
- **Cài như PWA / chạy đầy đủ:** đưa cả thư mục `web/` lên một host **https** (ví dụ GitHub
  Pages), mở link https đó, rồi Thêm vào Màn hình chính. Lúc này service worker mới đăng ký
  được — **service worker không chạy ở chế độ `file://`**.
- **HTML/JS thuần, không framework build**, không cần cài đặt, không thu thập dữ liệu,
  không nhúng key/secret, không ghi sẵn nội dung mẫu (nội dung demo trong ảnh chụp màn hình
  chỉ seed trong phiên chụp để kiểm chứng, không có trong app).

## Mới trong bản 1.2 — khung hình Office bản desktop Windows

Bản 1.1 bị chê đúng là “không giống Office tí nào” (chỉ là trang web di động chung chung:
tab viên thuốc, nút emoji, thẻ trắng xếp chồng). Bản 1.2 **đập giao diện cũ làm lại** theo
chuẩn duy nhất là **Office bản desktop Windows**:

- **Thanh tiêu đề màu theo app** (Word xanh dương #2B579A, Excel xanh lá #217346,
  PowerPoint cam đỏ #D24726) + **Quick Access Toolbar** (Lưu, Hoàn tác, Làm lại) + tên app,
  tên file, trạng thái Đã lưu/Chưa lưu.
- **Dải tab ribbon đủ tab theo từng món** (Word: Tệp, Trang đầu, Chèn, Thiết kế, Bố trí,
  Tham chiếu, Xem lại, Xem; Excel: Tệp, Trang đầu, Chèn, Bố trí Trang, Công thức, Dữ liệu,
  Xem lại, Xem; PowerPoint: Tệp, Trang đầu, Chèn, Thiết kế, Chuyển đổi, Trình chiếu, Xem
  lại, Xem) + **ribbon lệnh theo nhóm** (nhóm có tên, lệnh có icon ở trên + chữ ở dưới,
  nhóm cách nhau bằng vạch ngăn).
- **Khung Tệp (Backstage)** khi bấm tab Tệp: Mới / Mở / Lưu / Lưu thành / Thông tin, danh
  sách file kèm loại, ngày giờ, số slide.
- **Vùng làm việc nền xám + trang trắng bóng đổ**: Word = trang A4 có lề + thước ngang;
  Excel = hàng chữ cái cột + cột số hàng + **hộp tên ô + thanh công thức fx** + thẻ Sheet1
  dưới đáy; PowerPoint = **panel thu nhỏ slide bên trái + khung slide lớn ở giữa**.
- **Thanh trạng thái dưới đáy**: số trang/số từ (Word), ô đang chọn (Excel), số slide
  (PowerPoint) + thanh trượt thu phóng 50–200%.
- **Hành vi Windows**: Ctrl+B/I/U, Ctrl+S lưu im lặng; font mặc định **Carlito 11pt**
  (tương đương Calibri 11); file mặc định .docx/.xlsx/.pptx; **không còn emoji trong nút
  lệnh** (đã rà cả nút do mã sinh).
- **Thích ứng cảm ứng nhưng không mất dáng Office**: ribbon và thanh chuyển app cuộn
  ngang được; nút tối thiểu 48px; bottom sheet cho font/cỡ/màu/tiêu đề/mở file.
- Đối chiếu từng phần tử với Office Windows (cái nào giống, cái nào tương đương bằng cách
  khác, cái nào chưa làm và vì sao) xem file **`CHECKLIST_OFFICE_WINDOWS_1.2.md`** trong
  cùng thư mục — không có phần tử nào bị bỏ qua mà không ghi.

## Tính năng — đạt / chưa đạt

| Tính năng | Trạng thái |
|---|---|
| 1. Soạn văn bản (contenteditable + font, cỡ, đậm/nghiêng/gạch chân/gạch ngang, màu chữ, tô sáng, khoảng cách dòng, căn lề, gạch đầu dòng, tiêu đề; chèn bảng 3×3, ảnh, ngày, đường kẻ, ngắt trang, mục lục từ tiêu đề, nhận xét; lưu localStorage, mở lại đúng định dạng; xuất .txt và .docx chuẩn: styles + numbering thật, bảng thật, tô sáng thật) | ✅ Đạt (kiểm chứng vòng tròn LibreOffice, xem dưới) |
| 2. Bảng tính (lưới nhập số/chữ, `=SUM/=AVERAGE/=COUNT/=MAX/=MIN`, tham chiếu ô/vùng, phép tính + − × ÷ và ngoặc, tính lại khi sửa; thanh công thức fx + hộp tên ô; chèn/xóa hàng/cột; sắp xếp; lưu/mở lại; xuất .csv và .xlsx chuẩn có styles Carlito) | ✅ Đạt |
| 2a. Trình chiếu (panel thu nhỏ + khung slide hiện tại; tiêu đề + gạch đầu dòng; chèn ảnh nén tối đa 1280px; nền slide; cỡ 16:9/4:3; hiệu ứng mờ dần/trượt ngang; trình chiếu toàn màn hình từ đầu hoặc từ slide này, vuốt trái/phải hoặc nút lớn; lưu/mở lại; xuất .pptx chuẩn, font Carlito, nền slide ghi thật vào file) | ✅ Đạt |
| 3. PDF (chọn file, xem từng trang bằng pdf.js qua CDN; tô sáng, ghi chú, vẽ tay; chú thích lưu localStorage kèm tên file, mở lại thấy đủ) | ✅ Đạt, với giới hạn dưới đây |
| 3a. OCR (trang PDF thành ảnh hoặc ảnh riêng → văn bản sửa được, vie+eng) | ✅ Đạt bằng Tesseract.js (vie+eng) làm chính, OCR.space làm dự phòng |
| 3b. Chụp ảnh tạo PDF thật (camera/getUserMedia hoặc input capture, xem lại, lên/xuống/xóa trang, tạo PDF bằng JS thuần nhúng JPEG, lưu vào danh sách file để đồng bộ, mở lại bằng OCR; xref/startxref đã kiểm chứng đúng) | ✅ Đạt |
| 4. Dịch Việt ↔ Anh/Trung/Hàn/Nhật qua MyMemory (fetch, không key); dịch đoạn đã chọn rồi chèn ngay dưới đoạn gốc; mất mạng báo rõ, không bịa | ✅ Đạt |
| 5. Quản lý file (danh sách localStorage; nhập txt, docx văn bản thuần, csv, pdf; xuất/tải về; chia sẻ qua Web Share API khi trình duyệt hỗ trợ; xóa được cả bảng tính và PDF đang mở; tên trùng tự thêm (2), (3)…) | ✅ Đạt, với giới hạn docx dưới đây |
| 6. Font tiếng Việt (chỉ 6 font đã kiểm chứng, mặc định Carlito) | ✅ Đạt — xem mục Font |
| 7. Đăng nhập Google & Microsoft **ngay trong trang này** (mục “Đăng nhập & Đồng bộ”, 2 nút; đăng nhập xong lưu lên Drive/OneDrive và tải về mở lại ngay trong trang; trạng thái hiện trong danh sách file; đồng bộ giờ là file .docx/.xlsx/.pptx thật, không còn .txt/.csv) | ✅ Đạt về mặt mã (luồng OAuth chỉ chạy trọn vẹn qua https + redirect URI đã đăng ký — xem mục OAuth) |

## Giới hạn phải đọc kỹ

- **PDF bản 1.2 CHƯA ghi đè vào ruột file PDF gốc.** Chú thích chỉ lưu kèm tên file trong
  trình duyệt. Mở lại đúng file đó thì thấy đủ chú thích; file gốc giữ nguyên.
- **Nhập docx chỉ đọc văn bản thuần** (không giữ nguyên định dạng phức tạp, bảng, ảnh).
  Bản 1.2 đã sửa tận gốc để đọc được cả file docx kiểu **data-descriptor** (Word/LibreOffice
  hay ghi) bằng cách đọc thêm central directory khi đọc header thường thất bại — đã kiểm
  chứng bằng file tự tạo kiểu đó và file do LibreOffice lưu lại.
- **Chưa nhập ngược .xlsx/.pptx về sửa** (chỉ nhập txt, docx thuần, csv, pdf). Xuất thì đủ
  cả ba định dạng Office.
- **Một trang tính duy nhất (Sheet1)** trong mỗi bảng tính; chưa có nhiều trang tính,
  đóng băng ngăn, chia cửa sổ, in ấn, cộng tác thời gian thực, macro/Pivot/SmartArt.
  Xem đầy đủ lý do và cách thay thế trong `CHECKLIST_OFFICE_WINDOWS_1.2.md`.
- **File PDF gốc không lưu vào localStorage** (file có thể lớn). Chú thích lưu theo tên file;
  muốn xem lại chú thích thì chọn lại đúng file PDF đó. Riêng **PDF tạo bằng camera được lưu
  Blob vào IndexedDB** nên mở lại được từ danh sách file mà không cần chọn lại.
- **Đăng xuất Microsoft:** nút đăng xuất xóa token trong trang/phiên và gọi logout MSAL với
  đúng account (bản 1.1 có lỗi xóa account trước khi gọi logout nên lệnh logout không bao
  giờ chạy — bản 1.2 đã sửa); muốn thoát hẳn tài khoản Microsoft trên trình duyệt thì đăng
  xuất thêm ở tài khoản Microsoft nếu trình duyệt còn nhớ phiên.

## Font chữ — danh sách cuối cùng và cách kiểm chứng

Danh sách cuối cùng, **chỉ 6 font**, kèm font hệ thống dự phòng
(`-apple-system, "Segoe UI", sans-serif`):

1. **Carlito (mặc định)** — 2. Be Vietnam Pro — 3. Inter — 4. Roboto — 5. Noto Sans — 6. Noto Serif

Carlito là font mở **metric tương đương Calibri** (cùng độ rộng ký tự, là font mà
LibreOffice dùng để thay Calibri), **không nhúng font Microsoft** (bản quyền).

Cách kiểm chứng (bản 1.2, kiểm chứng thật, không chỉ đối chiếu metadata):

1. **Đối chiếu metadata Google Fonts:** Carlito và cả 5 font còn lại đều trả unicode-range
   phủ dấu tiếng Việt (riêng Carlito: khối vietnamese U+1EA0–1EF9 cùng các khối Latin mở rộng).
2. **Kiểm tra bằng fontTools trên file subset thật:** tải file subset vietnamese và latin của
   Carlito từ Google Fonts, hợp hai subset lại thì phủ đủ mọi ký tự trong đoạn kiểm chứng
   (“Tiếng Việt có dấu: à á ả ã ạ, ồ ộ ổ, ứ ự ử, đ, ê, ơ, ư — Bác sĩ Nguyễn Văn Đô đi học
   ở Hà Nội”), **không thiếu glyph nào** (lưu ý: ký tự ã thuộc subset latin, không thuộc
   khối vietnamese — kiểm tra từng subset riêng sẽ báo thiếu giả, phải hợp cả hai mới đúng).
3. **Đoạn kiểm chứng trong trang** (mục Cài đặt) kèm nút đổi qua **từng font** trong danh
   sách để nhìn bằng mắt; đổi font áp dụng ngay lên đoạn đang chọn và giữ đúng khi lưu/mở
   lại và khi xuất .docx (ghi đúng tên font vào `w:rFonts`, font mặc định ghi vào styles.xml).

Lưu ý trung thực: bước **nhìn bằng mắt cuối cùng cần Bùi kiểm tra trên máy thật** bằng các
nút đổi font trong trang; về mặt subset và fontTools thì cả 6 font đều phủ đầy đủ dấu tiếng
Việt nên không font nào vỡ dấu do thiếu glyph.

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
  Client ID vào mục Cài đặt (xem `Huong_dan_OAuth_Google_Microsoft.md` trong cùng thư mục).
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

- **Không tính năng nào bắt buộc chuột hay bàn phím:** không menu chỉ hiện khi rê chuột,
  không bấm chuột phải, không thao tác cần con trỏ chính xác từng pixel. Bàn phím chỉ dùng
  để gõ chữ (và phím tắt Ctrl+B/I/U/S khi có bàn phím); mọi nút chức năng đều bấm được bằng tay.
- **Vùng chạm tối thiểu 48px:** tất cả nút, ô nhập, ô chọn đều `min-height: 48px`; nút trong
  bottom sheet là 56px; nút trình chiếu 56px.
- **Ribbon và thanh chuyển app cuộn ngang được** trên điện thoại thay vì thu nhỏ đến mức
  không bấm được (đã chụp ảnh khổ 390px kiểm chứng).
- **Chọn font/cỡ/màu/tiêu đề/mở tài liệu bằng bảng chọn trượt từ dưới lên (bottom sheet)**
  thay vì menu thả xuống nhỏ; soạn tài liệu chạm vào là gõ.
- **Bảng tính:** chạm một lần để chọn ô, chạm hai lần liên tiếp để sửa ô, vuốt để cuộn.
- **PDF:** chụm hai ngón để phóng to/thu nhỏ, vẽ chú thích bằng ngón tay, chạm để thêm ghi chú.
- Lưu ý trung thực: rà soát này là rà logic mã + cấu trúc giao diện + ảnh chụp bằng trình
  duyệt máy tính (desktop 1366px và mobile 390px); Bùi vẫn nên bấm thử một vòng trên điện
  thoại thật để xác nhận cảm giác chạm cuối cùng (đặc biệt chụm phóng PDF và camera).

## Kiểm chứng bản 1.2 (giao diện + file + lỗi)

- `node --check app.js` đạt.
- **Harness Node/DOM giả (38 test, 38 pass, 0 fail):** xuất .docx/.xlsx/.pptx (xml well-formed,
  giữ đủ chữ tiếng Việt, định dạng, bảng, ảnh, font Carlito), tài liệu rỗng, công thức
  (SUM/AVERAGE/COUNT/MAX/MIN, phép tính, tham chiếu vòng tròn → #LỖI), PDF camera (header
  %PDF + %%EOF), trùng tên, xóa file đang mở, xóa bảng tính, đồng bộ khi chưa đăng nhập,
  dịch khi mất mạng (không bịa), camera bị từ chối quyền, đăng xuất Microsoft, chia sẻ lỗi
  (fallback tải file), trình chiếu kẹt đúng hai đầu, không xóa slide cuối, CSV có dấu phay
  trong ngoặc kép.
- **Vòng tròn LibreOffice 24.2 thật:** file do chính hàm export của app.js tạo →
  LibreOffice mở và chuyển sang PDF: .docx giữ đủ chữ tiếng Việt, đậm/nghiêng/gạch chân/
  gạch ngang, chữ đỏ, tô vàng, gạch đầu dòng có dấu •, bảng đủ 4 ô; .xlsx giữ đủ chữ và
  tính đúng SUM=30, A1+A2=30, AVERAGE=15 (chuyển sang CSV kiểm chứng đủ chữ dài, vì PDF
  của Calc cắt chữ theo độ rộng cột — đó là cách Calc in, không phải mất dữ liệu);
  .pptx mở được 2 trang, giữ đủ chữ và ảnh; tài liệu rỗng (.docx/.xlsx) mở được; tài liệu
  500 đoạn mở được đủ chữ; PDF camera 2 trang 400×300pt, xref/startxref kiểm chứng đúng
  từng offset.
- **Vòng ngược LibreOffice → app:** file .docx do LibreOffice lưu lại từ file app xuất
  (dạng deflate + central directory chuẩn) được hàm import của app đọc lại đủ chữ tiếng
  Việt và ô bảng; file .docx tự tạo kiểu data-descriptor (local header ghi size 0) cũng
  đọc được nhờ fallback central directory — đây là dạng file mà bản 1.1 đọc hỏng.
- **Kiểm chứng bằng mắt (bắt buộc của bản 1.2):** chụp bằng chrome-headless-shell local mở
  file://, điều khiển qua CDP, nội dung demo chỉ seed trong phiên chụp:
  `screenshots/v12-word.png`, `screenshots/v12-excel.png`, `screenshots/v12-powerpoint.png`,
  `screenshots/v12-backstage.png`, `screenshots/v12-word-mobile.png`. Soi ảnh đã bắt được
  3 lỗi và sửa tận gốc rồi chụp lại: (1) panel ribbon app cũ không ẩn khi đổi app → chồng
  panel; (2) lưới Excel bị ẩn do class `.sheet` của bottom-sheet đè lên `<table
  class="sheet">` (lỗi tiềm ẩn từ bản cũ, test logic không thấy); (3) chữ thumbnail slide
  dính nhau. Sau khi sửa, chụp lại và soi lại: cả 3 món đều ra dáng Office thật.
- **Kiểm chứng tương tác thật trong trình duyệt:** mở/đóng bottom sheet, chuyển tab ribbon
  (chỉ còn đúng 1 panel hiện), sửa thanh công thức rồi Enter là tính lại, Undo/Redo bảng
  tính, thêm/chọn slide, mở/đóng Backstage, thu phóng 150%, màu thanh tiêu đề Excel
  rgb(33, 115, 70) — tất cả pass (1 check “số 15” trả về dạng chuỗi "15" do displayCell luôn
  trả chuỗi — đó là cách hàm hiển thị làm việc, không phải lỗi tính toán).
- Giới hạn kiểm chứng: PowerPoint/WPS thật chưa kiểm chứng trực tiếp (không có máy đó ở
  đây) — Bùi mở thử file .pptx/.docx/.xlsx xuất từ trang trên máy thật để chốt cuối; VIP/
  đăng nhập OAuth thật, camera và chụm PDF trên điện thoại thật vẫn cần Bùi bấm thử.

## File trong thư mục web/

- `index.html` — toàn bộ giao diện khung Office Windows (title bar + ribbon + Backstage +
  status bar + 8 mục, tiếng Việt).
- `style.css` — giao diện Office, quy tắc cảm ứng 48px, bottom sheet, trang A4, lưới Excel,
  panel slide và màn hình trình chiếu.
- `app.js` — toàn bộ logic (soạn thảo, bảng tính, trình chiếu + xuất .pptx, PDF, OCR, camera
  tạo PDF, dịch, file, OAuth, khung Office).
- `manifest.json`, `sw.js` (cache `van-phong-bui-v3`), `icon.svg` — PWA (chỉ có tác dụng đầy
  đủ qua https).
- `Huong_dan_OAuth_Google_Microsoft.md` — hướng dẫn tự đăng ký app OAuth.
- `CHECKLIST_OFFICE_WINDOWS_1.2.md` — checklist đối chiếu từng phần tử Office Windows.
- `screenshots/` — ảnh chụp màn hình bản 1.2 để đối chiếu (không phải file của app, chỉ để
  kiểm chứng; khi đưa web lên host công khai có thể giữ hoặc bỏ thư mục này tùy Bùi).
- `README.md` — file này.
