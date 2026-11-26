# 04 — Chiến lược kiểm thử tự động

Mục tiêu: Claude **tự chạy** Word, áp lệnh lên bản sao fixture, đọc cấu trúc kết quả và so với golden — không cần người bấm.

## 1. Điều kiện môi trường (kiểm tra đầu phiên)
```powershell
# Word có cài không?
Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE' -ErrorAction SilentlyContinue
```
- Không có Word → chỉ chạy **test mức OOXML/logic thuần**; ghi rõ trong báo cáo "chưa kiểm chứng trong Word".
- Có Word + muốn build/chạy macro qua COM cần bật "Trust access to the VBA project object model" (`AccessVBOM`). Đây là **thay đổi cài đặt bảo mật → hỏi người dùng trước**, chỉ dùng trên máy dev, và ghi lại cách tắt.

### Môi trường đã quan sát (2026-11-23)
Word 16 đã cài (`C:\Program Files\Microsoft Office\Root\Office16`); `AccessVBOM` **chưa bật**; **không có** Python, Git, Node, dotnet → công cụ viết bằng PowerShell 5.1 + COM (đọc OOXML bằng `System.IO.Compression` + `XmlDocument`). Script phân tích cấu trúc đã dùng ở phiên này là mẫu cho `tools/inspect.ps1` (xuất style, `numbering`, page setup, field, định dạng trực tiếp — không xuất nội dung).
**AccessVBOM đã được bật ngày 2026-11-23 theo sự đồng ý của người dùng** (giá trị trước đó: chưa đặt). Chỉ ở HKCU, chỉ Word 16.0. Tắt lại khi xong M0/khi không dev:
```powershell
Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Security' -Name AccessVBOM
```
Dữ liệu thật: xem ADR-011 — fixture là bản tổng hợp; chạy trên văn bản thật chỉ ở "chế độ riêng" (bản sao thư mục tạm, báo cáo số liệu).

## 2. Các tầng test
| Tầng | Chạy gì | Cần Word? | Bắt lỗi gì |
|---|---|---|---|
| T1 Logic thuần | Hàm parse/đánh số (VBA không chạm Word, hoặc bản mô phỏng bằng PowerShell/Pester) | Không | Quy tắc số La Mã, nhận diện "Điều 1.", tính cấp |
| T2 Cấu trúc OOXML | Mở `.docx` kết quả như zip, đọc `word/document.xml`, `numbering.xml`, `styles.xml` | Không | Có `w:numPr`, style gắn list, TOC field (`w:fldChar`/`instrText TOC`), không còn "Điều 1" text gõ tay |
| T3 Tích hợp Word | PowerShell COM: mở bản sao, chạy macro, lưu, đọc `ListString`, `ListLevelNumber`, `Style` | **Có** | Hành vi thật: đánh số, TOC, undo, idempotent |
| T4 Thủ công có hướng dẫn | Danh sách kiểm 10 phút trên văn bản thật | Có | Giao diện Ribbon, cảm giác dùng |

## 3. Fixture (`tests/fixtures/`)
Đặt tên `fx-<nhóm>-<mô tả>.docx`. Bộ tối thiểu:
- `fx-hd-tho.docx` — hợp đồng chưa đánh số, 12 điều.
- `fx-hd-goTay.docx` — số gõ tay, có chỗ nhảy số 1,2,4.
- `fx-qd-chuong.docx` — quy định có Chương/Mục/Điều.
- `fx-cv-tho.docx` — công văn, lề/font lộn xộn.
- `fx-phuc-tap.docx` — có bảng, header/footer, footnote, ảnh, track changes, comment, section nhiều trang.
- `fx-trong.docx`, `fx-1doan.docx` — biên.
- `fx-dai-200trang.docx` — hiệu năng.
**Ẩn danh dữ liệu thật** trước khi đưa vào repo. Không commit văn bản mật.

## 4. Golden (`tests/golden/`)
Không so byte `.docx`. `tools/inspect.ps1` xuất JSON gồm, theo từng đoạn: `index`, `style`, `listString`, `listLevel`, `leftIndent`, `firstLineIndent`, `alignment`, `spaceBefore/After`, `lineSpacing(+rule)`, 30 ký tự đầu của text (để định vị), cộng cài đặt trang (lề, khổ giấy) và danh sách field. Test so JSON đầu ra với golden (cho phép dung sai số thực 0,05 pt).

## 5. Script chuẩn (Claude tạo trong `tools/` & `tests/` ở lần đầu)
- `tools/build.ps1` — tạo `dist/AddIn.dotm` bằng COM: mở Word ẩn, thêm module từ `src/` (`VBProject.VBComponents.Import`), nạp `customUI14.xml`, lưu `.dotm`.
- `tools/inspect.ps1 -Path x.docx` — dump JSON ở mục 4.
- `tests/run-tests.ps1` — với mỗi case: copy fixture vào thư mục tạm → mở bằng COM → nạp add-in → `Application.Run "Macro"` → lưu → inspect → so golden → in bảng PASS/FAIL, exit code ≠ 0 nếu có lỗi → **đóng Word và dọn tiến trình `WINWORD.EXE` mồ côi**.
- Luôn `try/finally`: `$doc.Close($false)`; `$word.Quit()`; `[Runtime.InteropServices.Marshal]::ReleaseComObject`.

## 6. Test bắt buộc cho mọi tính năng P0
1. **Đúng** (happy path) so golden.
2. **Idempotent**: chạy 2 lần → inspect y hệt lần 1.
3. **Undo**: chạy rồi `Application.Undo` 1 lần → inspect bằng bản gốc.
4. **Nội dung chữ không đổi**: so chuỗi text (đã bỏ số đầu dòng) trước/sau.
5. **Phức tạp**: `fx-phuc-tap.docx` không mất bảng/ảnh/footnote/comment; track changes không bị chấp nhận/loại bỏ ngầm.
6. **Biên**: tài liệu trống, 1 đoạn, vùng chọn rỗng, vùng chọn nằm trong bảng.
7. **Hiệu năng** (KB-E1): đo thời gian, in ra, fail nếu quá ngưỡng. Chính tả đo cả lần chạy lặp lại và **lần đầu trong phiên** (gọi `modSpell.ResetConfig` trước để nạp lại cấu hình + bảng bigram, như lần bấm đầu tiên của người dùng).

## 7. Báo cáo test
Mỗi lần chạy in:
```
KB-A1 happy         PASS
KB-A1 idempotent    PASS
KB-A1 undo          FAIL  đoạn 14: style 'VB-Dieu' != 'Normal'
Tổng: 21 PASS, 1 FAIL, 0 SKIP  (Word 16.0, 38.2s)
```
SKIP phải kèm lý do và không được dùng để che test lỗi.

## 8. Khi không có Word trên máy
Claude **không được** tuyên bố "đã test trên Word". Làm: T1+T2, nêu rõ phần chưa kiểm chứng, và cung cấp danh sách T4 cho người dùng tự chạy kèm tiêu chí đạt/trượt.

## 9. Bài học chạy test với Word qua COM (2026-11-23)
- Chạy script test trong **tiến trình con có timeout** (Start-Process powershell -File … ; WaitForExit), vì Word ẩn có thể treo (hộp thoại vô hình). Dọn Word mồ côi theo PID/thời điểm khởi tạo, không kill hàng loạt.
- Đường dẫn truyền cho Word phải là [string] thuần (không phải kết quả Join-Path thô).
- Application.Run chỉ nhận chuỗi/số và tìm macro theo tài liệu active → dùng modApi + Activate() tài liệu add-in (ADR-013).
- AccessVBOM có thể bị Word ghi đè về 0 khi thoát nếu đặt lúc Word đang chạy; đặt lại khi Word đóng.
- 	ests/run-tests.ps1 hiện chạy 27 ca (hello, build, macro, inspect, ribbon ×2, M1 ×21). Golden hiện nằm trong mã test (mảng ListString mong đợi); chuyển sang JSON ở M2 khi cần so cấu trúc lớn.
