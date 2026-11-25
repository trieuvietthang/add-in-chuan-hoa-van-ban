# 09 — Giải pháp & kiến trúc add-in

> Căn cứ: `08-PHAN-TICH-MAU-VAN-BAN.md` (cách bạn làm việc), ribbon của add-in tham chiếu (`00-DANH-GIA-TUONG-THICH.md` mục 5), môi trường máy bạn (Windows 11, **Word 16 đã cài**, chưa có Python/Node/Git; PowerShell 5.1).
> Tài liệu này là **thiết kế**; chưa có mã.

## 1. Ý tưởng cốt lõi: "Hồ sơ" (Profile) điều khiển mọi thứ
Một **Profile** = tập quy ước cho 1 loại văn bản. Add-in không hard-code số/định dạng; nó đọc Profile rồi *cài* và *áp dụng*.

```
Profile
 ├─ id, tên hiển thị (Hợp đồng / Quy định nội bộ / Công văn / …)
 ├─ page:     khổ giấy, lề, section, số trang (kiểu, bắt đầu từ trang mấy, vị trí)
 ├─ base:     font, cỡ, giãn dòng (multiple 1.15/1.2), giãn đoạn, thụt đầu dòng, căn lề
 ├─ levels[]: danh sách cấp, mỗi cấp có
 │     styleName, role (Chuong/Muc/Dieu/Khoan/Diem/Tiet/CongVan-I/…),
 │     numberFormat ("Điều %n.", "%1.%2", "%3)", "(%4)"), numberStyle (Arabic/Roman/letter),
 │     indent (số/ chữ/ tab), font override (đậm, hoa), outlineLevel, restartAfter
 ├─ toc:      danh sách style đưa vào mục lục + cấp, có số trang, dẫn chấm, style toc1..3
 ├─ checks:   bật/tắt từng luật kiểm tra, từ điển riêng
 └─ legacyAliases: tên style cũ nhận dạng được (01Dieu, 02Khoản, Heading 1…)
```
Profile mặc định (khởi tạo từ 3 mẫu của bạn — xem `06-NGHIEP-VU` mục 7):
- `HopDong`: Điều / Khoản (1.1) / Điểm (a) / Tiết ((i))
- `QuyDinh`: Chương / (Mục) / Điều / Khoản (1.) / Điểm (a) / Tiết ((i))
- `CongVan`: `I.` / `1.` / `1.1.` / `a)` / `(i)` / `-`
- Mỗi profile có biến thể chuẩn trang: `NĐ30` hoặc `ISO9001` (hoặc tùy chỉnh).

Lưu Profile ở **tệp JSON/INI UTF-8** cạnh add-in (`%APPDATA%\VBAddin\profiles\*.json`) → sửa được, xuất/nhập được (ý tưởng "Xuất/Nhập cấu hình" của repo tham chiếu). VBA không có parser JSON gốc → ADR-008: dùng định dạng đơn giản **INI-like** hoặc một parser JSON nhỏ tự viết (có test).

## 2. Các khối chức năng
```
┌─────────────── Ribbon "SOẠN THẢO VĂN BẢN" (customUI14) ──────────────┐
│ [Hồ sơ ▾] [Chuẩn hóa ▸] [Hoàn tác]                                    │
│ Cấu trúc: [Chương][Mục][Điều][Khoản][Điểm][Tiết] [↑Cấp][↓Cấp]        │
│ Mục lục:  [Chèn][Cập nhật]   Kiểm tra: [Dẫn chiếu][Chính tả][Cấu trúc]│
│ Cấu hình: [Sửa hồ sơ][Xuất][Nhập]                                     │
└───────────────────────────────────────────────────────────────────────┘
```
| Module (`src/…`) | Trách nhiệm | Chạm Word? |
|---|---|---|
| `core/modProfile` | đọc/ghi/validate Profile | không |
| `core/modStyles` | tạo/sửa style theo Profile (`EnsureStyles`) | có |
| `core/modLists` | tạo/sửa ListTemplate đa cấp gắn style (`EnsureListTemplate`) | có |
| `core/modStructure` | **mô hình cấu trúc**: duyệt đoạn → vai trò → địa chỉ ("Điều 5 khoản 2 điểm a") | đọc |
| `features/modDetect` | nhận diện vai trò từ style cũ/tiền tố gõ tay/heuristic + độ tin cậy | đọc |
| `features/modApply` | áp style, xóa tiền tố gõ tay, đổi cấp, khởi động lại số | ghi |
| `features/modNormalize` | pipeline chuẩn hóa (mục 3) | ghi |
| `features/modToc` | TOC field + style toc | ghi |
| `features/modRefs` | trích dẫn chiếu, kiểm tra, chuyển REF field, sửa theo bản đồ đổi số | đọc/ghi |
| `features/modSpell` | chính tả offline (mục 5) | đọc |
| `features/modReport` | khung báo cáo (UserForm) liệt kê phát hiện, bấm nhảy tới vị trí, sửa nhanh | UI |
| `core/modLog`, `modStrings`, `modUndo` | log, chuỗi tiếng Việt, bao UndoRecord/khôi phục | — |
| `ui/` | ribbon XML + callback mỏng | — |

Nguyên tắc: logic **thuần** (`modProfile`, regex dẫn chiếu, phân loại âm tiết, tính bản đồ đổi số) tách riêng để test không cần Word.

## 3. Pipeline "Chuẩn hóa" (một nút, idempotent, có xem trước)
```
0 Tiền kiểm   : tài liệu bị bảo vệ? track changes? có field TOC? → cảnh báo/ chế độ an toàn
1 Nhận diện   : mỗi đoạn → vai trò + độ tin cậy  ⇒ "Bảng kế hoạch cấu trúc"
2 Xem trước   : hộp liệt kê (đoạn · vai trò hiện tại → đề xuất · tin cậy); người dùng sửa/ bỏ chọn
3 Cài hồ sơ   : EnsureStyles + EnsureListTemplate (tạo nếu thiếu, sửa nếu lệch — không nhân bản list)
4 Áp cấu trúc : gán style; xóa tiền tố số gõ tay; chuyển numPr trực tiếp → style; lưu bản đồ "số cũ → số mới"
5 Định dạng   : Normal + style hồ sơ (font/giãn/thụt/căn lề); xóa định dạng ký tự trực tiếp thừa (font, cỡ, màu), GIỮ đậm/nghiêng/gạch chân/highlight có chủ đích (tùy chọn)
6 Trang       : A4, lề, header/footer, số trang theo hồ sơ; cảnh báo section thừa
7 Dẫn chiếu   : sửa dẫn chiếu gõ tay theo bản đồ đổi số (nếu bật) hoặc chuyển REF
8 Mục lục     : chèn/cập nhật nếu hồ sơ yêu cầu
9 Báo cáo     : tóm tắt đã đổi gì; danh sách cần người duyệt
```
Quy tắc an toàn: không đổi chữ nội dung (trừ xóa tiền tố số gõ tay đã xác nhận); không đụng bảng/ảnh/footnote/comment/track changes; một `UndoRecord` cho toàn pipeline.

### Nhận diện vai trò (bước 1) — tín hiệu theo thứ tự tin cậy
1. Style cũ có trong `legacyAliases` (01Dieu, 02Khoản, Heading 1…) → tin cậy cao.
2. Đoạn có list thật (`ListString`, `ListLevelNumber`) → ánh xạ theo kiểu số (`Điều N.`→Điều, `a)`→Điểm…).
3. Tiền tố gõ tay bằng regex: `^(PHẦN|Phần)\s+[IVX\d]+`, `^(CHƯƠNG|Chương)\s+[IVXLC\d]+`, `^Mục\s+\d+`, `^Điều\s+\d+[.:]?`, `^\d+\.\d+\.?\s`, `^\d+\.\s`, `^[a-zđ]\)\s`, `^\(?[ivx]+\)\s`, `^[-–•]\s`.
4. Heuristic (tin cậy thấp, luôn cần duyệt): ALL CAPS + ngắn + đậm → Chương/tiêu đề; đoạn đậm kết thúc không dấu chấm → tên Điều.
5. Ngữ cảnh: chuỗi số phải hợp lệ (Điều 1,2,3… ; sau Điều mới có Khoản) — dùng để nâng/giảm tin cậy và phát hiện nhảy số.

## 4. Thiết kế đánh số (điểm kỹ thuật khó nhất)
- Mỗi Profile → **một** `ListTemplate` đa cấp (`OutlineNumbered:=True`) tên cố định `VB-<profile>`, mỗi cấp `LinkedStyle` = style của cấp đó. Style **đặt thụt/Tab** ở list level (`NumberPosition`, `TextPosition`, `TabPosition`) — không dùng thụt trong style để tránh xung đột.
- Cấp "Điều" trong QuyDinh đếm **liên tục xuyên Chương** (`ResetOnHigher` của cấp Điều = Mục/Chương? → cấu hình `restartAfter: none|Chuong|Muc`). Khoản reset sau Điều.
- Chuyển profile trong cùng văn bản: liên kết lại style → list level (một style chỉ liên kết một cấp tại một thời điểm); thao tác bọc Undo.
- **Tránh sinh list mồ côi**: luôn tìm `ListTemplate` theo tên trước khi tạo; không `ListGalleries` mỗi lần.
- Tên có dấu trong số ("Điều %n.") dựng bằng `modStrings` (`ChrW`).
- Công văn dùng cùng cơ chế, cấp: `I.` (La Mã hoa) → `1.` → `1.1.` → `a)` → `(i)` → `-`.

## 5. Kiểm tra
### 5.1 Dẫn chiếu (ưu tiên cao — rủi ro pháp lý)
1. Dựng **chỉ mục địa chỉ** từ `ListString`/cấp: `Điều 5`, `Điều 5.2`, `Điều 5 khoản 2 điểm a`, `Phụ lục 01`.
2. Trích dẫn chiếu bằng regex (nhóm nội bộ: `Điều N`, `khoản N`, `điểm x`, `Mục N`, `Chương N`, `Phụ lục N`, "Điều này", "khoản này"). **Loại trừ dẫn chiếu ngoài** khi theo sau là `của (Luật|Bộ luật|Nghị định|Thông tư|Quyết định|Nghị quyết|…)` hoặc có số hiệu văn bản.
3. Luật: R1 đích không tồn tại; R2 tên điều trong ngoặc không khớp tiêu đề thật; R3 tự dẫn chiếu vô nghĩa; R4 dẫn chiếu gõ tay (không phải REF) — khuyến nghị chuyển; R5 dẫn chiếu "ra ngoài phạm vi" (khoản không có trong Điều); R6 dẫn chiếu tới phần đã xóa (sau đổi số).
4. Chuyển REF: tạo bookmark ẩn `_RefVB_<id>` trên đoạn đích; chèn `REF _RefVB_<id> \r \h` (số kèm "Điều"); cập nhật bằng `Fields.Update`. **Cần spike** kiểm tra Word trả về gì với định dạng "Điều %n." (kể cả bỏ dấu chấm cuối).
5. Bản đồ đổi số khi chuẩn hóa: `số cũ → số mới` theo danh tính đoạn; dùng để tự sửa dẫn chiếu gõ tay (có xem trước).

### 5.2 Chính tả (offline, tầng)
Word thường **không có** công cụ kiểm tra tiếng Việt → không dựa vào `Document.SpellingErrors`. (Spike S5 xác nhận trên máy bạn.) Thiết kế 4 tầng, tầng sau không cần mạng:
1. **Kiểu chữ/dấu câu** (chắc chắn): hai dấu cách, dấu cách trước `, ; : . ! ?`, thiếu dấu cách sau, lặp từ ("của của"), ngoặc/ nháy không cân, thứ tự dấu, "Điều" viết thường đầu câu, số/ngày định dạng.
2. **Từ điển lỗi thường gặp** (người dùng sửa được): "xử lí/xử lý", "hoà/hòa" — kiểu bỏ dấu (`hoá/hóa`) theo lựa chọn hồ sơ; viết tắt không nhất quán.
3. **Kiểm tra âm tiết hợp lệ**: tách từ theo khoảng trắng; mỗi âm tiết kiểm tra tập âm tiết tiếng Việt hợp lệ (nạp từ tệp UTF-8). Bắt lỗi gõ sai không thành âm tiết ("hợp đồg", "thanh tóan"). Không bắt lỗi "đúng âm tiết nhưng sai từ" (ví dụ hỏi/ngã) → ghi rõ giới hạn.
4. **Thuật ngữ & nhất quán**: các thuật ngữ định nghĩa ("Bên A", "Hợp Đồng") dùng nhất quán hoa/thường; Bên A/Bên B không lẫn; số tiền/bằng chữ khác nhau (tùy chọn P2).
Soát nghĩa/ngữ pháp sâu: ngoài phạm vi add-in (offline); nếu muốn, quy trình bên ngoài: xuất báo cáo → nhờ Claude đọc văn bản (có sự đồng ý của bạn về dữ liệu).

## 6. Mục lục
- Chèn field `TOC \h \z \t "<style1>,1,<style2>,2,…" \u` (liệt kê style của hồ sơ + outline) — dấu phân cách danh sách theo `Application.International(wdListSeparator)` (có thể là `;`).
- Style `VB-Toc1..3` (hoặc `TOC 1..3` ghi đè) theo hồ sơ: thụt, dẫn chấm, TNR; **không** để mục lục tự rơi vào mục lục (tiêu đề "MỤC LỤC" dùng style không có outline).
- Hợp đồng dài: chỉ Điều; Quy định: Chương → Điều; tuỳ chọn thêm Phụ lục; tuỳ chọn không số trang (`\n`).
- Cập nhật: `TablesOfContents(1).Update`; cảnh báo khi trường bị khóa/ trống.

## 7. Hoàn tác
UndoRecord bao cả pipeline (đạt Ctrl+Z một lần). Rủi ro: một số thao tác (tạo ListTemplate, field) có thể không vào nhóm undo hoặc UndoRecord bị cắt khi Word crash. Phương án dự phòng (spike S6): trước khi chạy **Chuẩn hóa**, nhân bản nội dung vào tài liệu ẩn tạm (hoặc hỏi người dùng lưu bản sao) → nút **Hoàn tác** của add-in (ý tưởng nút "Hoàn tác" ở repo tham chiếu).

## 8. Tận dụng kiến trúc từ repo tham chiếu (chỉ khái niệm; không sao chép mã)
| Ý tưởng quan sát được ở repo | Dùng cho dự án này |
|---|---|
| Ribbon `customUI14.xml` kèm tab riêng + nhóm ở tab Home | Giữ: tab "SOẠN THẢO VĂN BẢN" + nhóm nhỏ ở Home cho "Chuẩn hóa"/"Hoàn tác" |
| Dropdown **Khối** → bộ quy tắc | Đổi thành dropdown **Hồ sơ** (Hợp đồng/Quy định/Công văn) |
| Nút **Chuẩn Hóa** một chạm + **Hoàn tác** riêng | Giữ nguyên ý tưởng |
| **Tùy chọn thể thức** chỉ nhận giá trị trong khoảng cho phép | Hồ sơ có `min/max` cho lề, cỡ chữ, giãn dòng; ngoài khoảng → cảnh báo |
| **Xuất/Nhập cấu hình** | Giữ (profile JSON/INI) |
| Danh sách đơn vị ban hành, tạo văn bản mẫu | **Không làm giai đoạn đầu** (bạn không cần) |
| Nút Cập nhật qua mạng + bản quyền | **Không làm** (ADR-005) — phân phối thủ công/ GPO |
| Hai bản `.dotm` Win/Mac | Bỏ Mac (ngoài phạm vi) |
| `.dotm` chứa ribbon XML trong package | Giữ: `build.ps1` phải **chèn `customUI/customUI14.xml` vào gói zip** sau khi Word lưu `.dotm` (VBA không tự tạo ribbon) |

## 9. Triển khai & môi trường
- Tệp `.dotm` đặt ở `%APPDATA%\Microsoft\Word\STARTUP` (tải tự động) hoặc Trusted Location; ký số VBA nếu có chứng chỉ (nội bộ/ self-signed cho máy bạn).
- Dev loop cần COM nạp module VBA → bật **AccessVBOM** (HKCU) trên **máy dev** (đang chưa bật) — thay đổi cài đặt bảo mật, **cần bạn đồng ý**, và có nút tắt lại sau khi xong. Phương án không cần bật: bạn nhập module bằng tay một lần (chậm, không tự test được).
- Không có Python/Git/Node: toàn bộ công cụ build/test viết bằng **PowerShell 5.1** + COM.
- Văn bản thật của bạn là bí mật: **không đưa vào repo**; fixture là bản **tổng hợp** có cùng cấu trúc (script tạo), chỉ dùng văn bản thật trong "chế độ riêng" trên bản sao ở thư mục tạm, kết quả báo cáo chỉ gồm số liệu cấu trúc.

## 10. Rủi ro kỹ thuật & giảm thiểu
| Rủi ro | Mức | Giảm thiểu |
|---|---|---|
| Liên kết style–list level sai/ lệch khi tài liệu đã có list lạ | Cao | Chỉ tạo list của hồ sơ; chuyển numPr trực tiếp → style; test với F1/F3-dạng |
| Xóa định dạng ký tự làm mất đậm/nghiêng có chủ đích | Cao | Chỉ xóa font/cỡ/màu; giữ B/I/U/highlight; chế độ "mạnh" tùy chọn + xem trước |
| REF trả chuỗi không như mong muốn | Trung bình | Spike S3; phương án B: chèn văn bản + bookmark và "cập nhật dẫn chiếu" bằng macro |
| UndoRecord không bao trọn | Trung bình | Spike S6; backup tạm |
| Tài liệu 200+ trang chậm | Trung bình | Duyệt `For Each`; tắt ScreenUpdating; thanh tiến độ |
| Macro bị chặn trên máy công ty | Trung bình | Ký số; Trusted Location; ADR để cân nhắc COM add-in/Office.js sau |
| Chính tả sai âm tiết nhưng đúng từ không bắt được | Thấp | Nói rõ giới hạn; từ điển riêng |
| Word tiếng Việt vs tiếng Anh (tên style, list separator) | Trung bình | Không dùng tên style bản địa; đọc `International()` |
