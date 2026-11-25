# 01 — Kịch bản người dùng & tiêu chí chấp nhận

Mỗi kịch bản có ID (`KB-xx`) — test, commit và issue phải trích ID này. Định dạng: **Cho** (tiền điều kiện) / **Khi** / **Thì**.
Ưu tiên: P0 = bắt buộc bản đầu; P1 = bản kế; P2 = sau.

## Nhóm A — Đánh số điều khoản

### KB-A1 (P0) Áp bộ đánh số "Hợp đồng"
- **Cho** văn bản có các đoạn tiêu đề điều khoản chưa đánh số.
- **Khi** người dùng chọn các đoạn (hoặc cả văn bản) và bấm "Đánh số → Hợp đồng".
- **Thì** cấp 1 = `Điều 1.`, cấp 2 = `1.1.`, cấp 3 = `a)`, cấp 4 = `-`; số là list thật (Word tự đếm); thụt dòng theo `06-NGHIEP-VU`.
- Chấp nhận: chèn thêm một Điều vào giữa → toàn bộ số sau tự cập nhật; không có số nào gõ tay.

### KB-A2 (P0) Áp bộ đánh số "Quy định nội bộ"
- Cấp: `Chương I` (số La Mã) → `Mục 1` → `Điều 1.` → `1.` → `a)`. **Số Điều đếm liên tục xuyên Chương** (không reset khi sang Chương mới) — cấu hình được.

### KB-A3 (P0) Tăng/giảm cấp nhanh
- **Khi** bấm Tab/Shift+Tab hoặc nút "Tăng cấp/Giảm cấp" (cũng có phím tắt).
- **Thì** đoạn đổi cấp, số tự đổi đúng kiểu; không đổi nội dung chữ.

### KB-A4 (P0) Chuyển số gõ tay thành list thật
- **Cho** văn bản có "Điều 1.", "1.1", "a)" gõ bằng tay.
- **Khi** chọn "Chuyển số gõ tay → đánh số tự động".
- **Thì** phần số gõ tay bị xóa khỏi text, thay bằng list thật cùng cấp; văn bản sai thứ tự (nhảy số 1,2,4) được đánh lại đúng. Có bước **xem trước** danh sách thay đổi.

### KB-A5 (P1) Khởi động lại / tiếp tục số
- Bắt đầu lại từ số N; tiếp tục list trước đó (hợp đồng có phụ lục đánh số riêng: `Điều 1` của Phụ lục 01 không dính số với thân hợp đồng).

### KB-A6 (P1) Tham chiếu chéo
- Chèn "Điều X" / "khoản Y Điều X" bằng REF field; khi số đổi, bấm "Cập nhật tham chiếu" → đúng. Báo cáo tham chiếu bị gãy.

## Nhóm B — Định dạng chuẩn

### KB-B1 (P0) Áp thể thức **Công văn**
- Theo `06-NGHIEP-VU` (NĐ 30/2020): khổ A4, lề, font Times New Roman cỡ 13–14, đoạn văn thụt đầu dòng, giãn dòng, số trang từ trang 2 ở giữa đầu trang.
- Chấp nhận: gọi 2 lần ra cùng kết quả (idempotent); không phá bảng/ảnh.

### KB-B2 (P0) Áp thể thức **Hợp đồng** và **Quy định nội bộ**
- Hồ sơ định dạng (profile) riêng, lưu thành Style của Word (`VB-Normal`, `VB-Dieu`, `VB-Khoan`, `VB-Diem`, `VB-Tieude`…).

### KB-B3 (P0) Thụt dòng & giãn đoạn
- Nút "Thụt đầu dòng 1,27 cm", "Thụt cấp", "Giãn dòng 1,15/1,5/Exactly". Áp trên vùng chọn; mặc định áp cho cả văn bản khi không chọn gì (có xác nhận).

### KB-B4 (P1) Chuẩn hóa văn bản dán
- Xóa khoảng trắng đôi, dòng trống thừa, dấu cách trước dấu câu; chuẩn hóa dấu ngoặc kép; thay Tab/space dùng để thụt dòng bằng thụt dòng thật. Có **xem trước** + hoàn tác.

### KB-B5 (P1) Tiêu ngữ, quốc hiệu, bố cục bảng đầu công văn
- Chèn khối quốc hiệu–tiêu ngữ + tên cơ quan (bảng 2 cột không viền) theo mẫu.

### KB-B6 (P1) Khối chữ ký, nơi nhận
- Chèn mẫu "Nơi nhận" + chức vụ/chữ ký.

## Nhóm C — Mục lục

### KB-C1 (P0) Chèn mục lục
- **Khi** bấm "Mục lục" tại con trỏ.
- **Thì** chèn TOC field (cấp 1–3 hoặc tùy chọn), có dẫn chấm và số trang; "Cập nhật mục lục" làm tươi field.
- Với hợp đồng: mục lục lấy `Điều N. Tên điều`; với quy định: gồm cả `Chương`.

### KB-C2 (P1) Mục lục chỉ liệt kê Điều / không số trang / tạo ở đầu văn bản.

## Nhóm D — Kiểm tra & báo cáo

### KB-D1 (P1) Kiểm tra cấu trúc
- Báo: nhảy số, trùng số, Điều không tên, cấp bị bỏ (Điều → a) không qua Khoản), tham chiếu gãy, định dạng lệch chuẩn. Hiển thị danh sách, bấm nhảy tới vị trí.

## Nhóm E — Phi chức năng

| ID | Yêu cầu | Ngưỡng |
|---|---|---|
| KB-E1 | Hiệu năng | 200 trang, ≤ 3.000 đoạn: mỗi lệnh ≤ 5 giây; không treo UI (bật `ScreenUpdating=False`, khôi phục khi lỗi) |
| KB-E2 | Hoàn tác | 1 lệnh = 1 lần Ctrl+Z |
| KB-E3 | An toàn | Không sửa nội dung chữ; không mất ảnh, bảng, footnote, comment, track changes |
| KB-E4 | Tương thích | Word 2016/2019/2021/365 Windows; ngôn ngữ giao diện Anh/Việt (không phụ thuộc tên Style bản địa — dùng `wdStyle*` hoặc style tự tạo) |
| KB-E5 | Bảo mật | Không mạng, không tự cập nhật, mã nguồn có thể ký số |
| KB-E6 | Lỗi | Mọi lỗi → hộp thoại tiếng Việt dễ hiểu + ghi log; trạng thái Word luôn được khôi phục |

## Bổ sung 2026-11-23 (sau phân tích 3 mẫu — xem `08`, `09`)

**Thay đổi so với bản đầu:** tên style chính thức theo thói quen của bạn (`00Chuong, 01Dieu, 02Khoan, 03Diem, 04Tiet`, thêm `00Phan/00Muc` nếu cần) thay cho tiền tố `VB-`; nhận diện được style cũ (`02Khoản` có dấu, `Heading 1…`) qua `legacyAliases`. KB-A1/A2 áp dụng với các style này. Số Khoản/cấp Điều cấu hình theo **Hồ sơ**.

### KB-A7 (P0) Hồ sơ **Công văn** — đánh số nhiều cấp
- Cấp: `I.` → `1.` → `1.1.` → `a)` → `(i)` → `-`; áp bằng nút cấp/ tăng–giảm cấp; số tự đếm, đúng thụt; dãy ký hiệu chỉnh trong hồ sơ.

### KB-D2 (P0) Kiểm tra dẫn chiếu sai
- **Cho** văn bản có "Điều X/khoản Y/điểm z" gõ tay hoặc REF.
- **Khi** bấm "Kiểm tra dẫn chiếu".
- **Thì** báo: đích không tồn tại; tên điều trong ngoặc không khớp tiêu đề; khoản/điểm không thuộc Điều; dẫn chiếu gõ tay (đề nghị chuyển REF); bỏ qua dẫn chiếu tới văn bản ngoài ("Điều 5 của Luật …"). Danh sách bấm nhảy tới vị trí; có nút "Chuyển REF".

### KB-D3 (P1) Kiểm tra chính tả/ kiểu chữ (offline)
- 4 tầng (kiểu chữ-dấu câu; từ điển lỗi thường gặp; âm tiết hợp lệ; nhất quán thuật ngữ). Không đổi văn bản tự động, chỉ báo/đề xuất. Nói rõ giới hạn: không bắt lỗi "đúng âm tiết, sai từ".
- **Chi tiết M5 (chốt 2026-11-25):** nút "Kiểm tra chính tả" + "Lỗi tiếp theo" (như dẫn chiếu), chỉ báo, không sửa chữ. Mã lỗi và tiêu chí chấp nhận (test trên bản tổng hợp):
  | Mã | Lỗi | Ví dụ bắt | Không báo (chống báo nhầm) |
  |---|---|---|---|
  | S1 | hai dấu cách trở lên giữa chữ | `A  B` | thụt đầu dòng/tab, dấu cách đầu-cuối đoạn |
  | S2 | dấu cách trước `, ; : ! ? )` hoặc `.` cuối câu | `hợp đồng ,` | `...` |
  | S3 | thiếu dấu cách sau `, ;` hoặc sau `.` giữa chữ thường–chữ HOA | `bên A,bên B`, `xong.Tiếp` | `1,5`, `12:30`, `TP.HCM`, địa chỉ web/email |
  | S4 | lặp từ liền kề | `của của` | số, từ láy trong `config/spell.vi.ini` (`dần dần`) |
  | S5 | ngoặc/nháy không cân trong đoạn | `(khoản 2` | số thứ tự gõ tay đầu đoạn `a)`, `1)` |
  | S6 | "điều N" viết thường đầu câu | `. điều 5 quy định` | "điều 5" giữa câu |
  | W1 | cặp sai→đúng trong từ điển người dùng | `xem sét` → `xem xét` | — |
  | W2 | cùng một từ viết hai kiểu (nhóm biến thể) | `quy định` và `qui định` cùng văn bản | chỉ có một kiểu |
  | W3 | âm tiết không hợp lệ tiếng Việt (âm đầu–vần–âm cuối–dấu thanh–vị trí dấu) | `đồg`, `nghành`, `tóan` | từ viết hoa toàn bộ, có chữ số, có f/j/w/z, token toàn ASCII, từ trong danh sách bỏ qua |
  - Quét cả bảng, chú thích; bỏ qua mã trường (field code) và mục lục; chỉ đọc văn bản (nhiều lần cho cùng kết quả). Từ điển/danh sách bỏ qua do người dùng sửa trong `config/spell.vi.ini`.
  - Báo kết quả dạng tóm tắt theo loại + nhảy tới từng lỗi; 200 trang ≤ 5 giây; không mất Undo (lệnh không ghi gì vào văn bản).
  - Giới hạn nói rõ: không bắt lỗi "đúng âm tiết, sai từ" (hỏi/ngã…), từ nước ngoài không dấu, tên riêng lạ.

### KB-D4 (P1) Kiểm tra cấu trúc
- Nhảy số, trùng số, Điều không tên, Điểm không nằm dưới Khoản, section thừa, list mồ côi, đoạn "điều khoản" bỏ qua style (numPr trực tiếp).

### KB-F1 (P0) Chuẩn hóa một chạm theo Hồ sơ (+ xem trước + Hoàn tác)
- **Khi** bấm "Chuẩn hóa" với hồ sơ đã chọn (Hợp đồng/Quy định/Công văn × NĐ30/ISO9001/tùy chỉnh).
- **Thì** chạy pipeline `09` mục 3: nhận diện → xem trước → cài hồ sơ → áp style → trang/phông/giãn → (dẫn chiếu) → (mục lục) → báo cáo. Kết quả: đúng khổ A4 + lề; font/cỡ/giãn đồng nhất (ví dụ TNR 12–13, multiple 1.15/1.2); không còn định dạng ký tự thừa (font/cỡ/màu); chữ nội dung không đổi; một Ctrl+Z hoặc nút "Hoàn tác" khôi phục.
- Nghiệm thu: chạy trên bản sao 3 mẫu thật (cấu trúc: `08` mục 1) — các vấn đề P1–P12 được xử lý hoặc báo cáo.

### KB-G1 (P1) Sửa hồ sơ, xuất/nhập cấu hình
- Chỉnh thông số trong khoảng cho phép (lề, cỡ chữ, giãn dòng, ký hiệu cấp); xuất/nhập tệp hồ sơ giữa các máy.

### Sắp xếp lại ưu tiên (theo giá trị cho bạn)
P0: A1–A4, A7, B1–B3, C1, D2, F1, E1–E6 · P1: A5, A6, B4–B6, C2, D1(=D4), D3, G1 · P2: tạo văn bản mẫu, tự cập nhật (loại bỏ).

## Quy tắc quản lý kịch bản
- Thêm kịch bản mới ở cuối nhóm, ID không tái sử dụng.
- Kịch bản P0 chưa có test tự động = chưa "xong".
- Khi làm rõ nghiệp vụ với người dùng → cập nhật tại đây và `06-NGHIEP-VU`.
