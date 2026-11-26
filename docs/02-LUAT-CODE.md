# 02 — Luật viết code (VBA cho Word)

## 1. Nguyên tắc chung
1. **Source = text trong `src/`**. Không sửa `dist/*.dotm` trực tiếp. Sửa trong VBE → chạy `tools/export.ps1` để đưa về `src/` rồi commit.
2. Một việc một hàm; hàm ≤ 40 dòng; module ≤ 400 dòng. Tách logic **thuần** (tính toán, parse) khỏi logic **chạm Word** để test được.
3. `Option Explicit` ở mọi module. Không `On Error Resume Next` trần trụi (xem mục 5).
4. Tên: module `modXxx` / lớp `clsXxx` / form `frmXxx`; hàm công khai `PascalCase`; biến `camelCase` có tiền tố kiểu (`rng`, `para`, `lt`, `s`, `n`, `b`). Hằng `UPPER_SNAKE`.
5. Chú thích tiếng Việt có dấu ở **comment**; ngắn, nêu *vì sao*, không nhắc lại *cái gì*.

## 2. Tiếng Việt trong VBA
- File `.bas/.cls` lưu ANSI (Windows-1252) nên **không viết chữ có dấu trong chuỗi**. Mọi chuỗi hiển thị đặt trong `modStrings` dưới dạng hàm trả về chuỗi dựng bằng `ChrW(&H…)` (có bảng tra trong comment), hoặc đọc từ `resources/strings.vi.txt` UTF-8 qua `ADODB.Stream` (chọn một, ghi ADR).
- Test bắt buộc: "Điều", "Chương", "Mục", "Quốc hiệu" hiển thị đúng trong Word thật.

## 3. Thao tác với Word
- **Không dùng `Selection`** để sửa. Dùng `Range`, `Paragraph`, `ListFormat`. Cho phép đọc `Selection.Range` ở đầu lệnh để biết vùng làm việc.
- **Mọi lệnh ghi** bọc bởi:
  ```vb
  Dim rec As UndoRecord: Set rec = Application.UndoRecord
  rec.StartCustomRecord "Ten lenh"
  On Error GoTo EH
  ' ... làm việc ...
  Cleanup: rec.EndCustomRecord: Exit Sub
  EH: ' ghi log, thông báo, rồi GoTo Cleanup
  ```
- Tắt `Application.ScreenUpdating` trong lệnh dài và **luôn khôi phục** ở nhánh `Cleanup`, kể cả khi lỗi.
- Không dùng chỉ số tuyệt đối (`Paragraphs(37)`) qua nhiều thao tác làm đổi số lượng đoạn; duyệt ngược hoặc giữ tham chiếu `Range`.
- Không hard-code tên Style tiếng Anh/Việt. Dùng hằng `wdStyle…` cho Style dựng sẵn; style của dự án tạo bằng `modStyles.EnsureStyles` với tiền tố `VB-`.
- Style trước, định dạng trực tiếp sau cùng: **ưu tiên đặt định dạng trong Style**, chỉ dùng direct formatting khi bắt buộc. Cấm xóa/ghi đè định dạng ký tự (đậm/nghiêng) của người dùng khi áp thể thức đoạn, trừ khi lệnh nói rõ.

## 4. Đánh số (quan trọng nhất)
- Mỗi bộ đánh số = 1 `ListTemplate` đa cấp (`ListGalleries(wdOutlineNumberGallery)` hoặc `ActiveDocument.ListTemplates.Add(OutlineNumbered:=True, Name:="VB-HopDong")`) **gắn LinkedStyle** cho từng cấp (`ListLevels(i).LinkedStyle = "VB-Dieu"`…).
- Cấu hình từng cấp: `NumberFormat`, `NumberStyle`, `NumberPosition`, `TextPosition`, `TabPosition`, `TrailingCharacter`, `ResetOnHigher`, `StartAt`, `Font` (chỉ khi cần).
- "Điều 1." = `NumberFormat = "Đi" & ChrW(&H1EC1) & "u %1."` (dựng bằng `modStrings`), `NumberStyle = wdListNumberStyleArabic`. La Mã chương: `wdListNumberStyleUppercaseRoman`.
- Đổi cấp bằng cách đổi **Style** của đoạn (không đổi `ListLevelNumber` thủ công) để không lệch.
- Tham chiếu chéo: `InsertCrossReference ReferenceType:=wdRefTypeNumberedItem` — không tự gõ số.

## 5. Xử lý lỗi & log
- Mẫu chuẩn: `On Error GoTo EH` + nhánh `EH` gọi `modLog.LogError "Tên.Hàm", Err`; nhánh `Cleanup` khôi phục trạng thái. `On Error Resume Next` chỉ cho **1 câu lệnh** có chủ đích, kèm comment và `Err.Clear`.
- `modLog`: ghi file `%TEMP%\VBAddin\log.txt` (UTF-8 không BOM qua `ADODB.Stream` hoặc ANSI không dấu), xoay vòng 1 MB. Không ghi nội dung văn bản người dùng vào log (riêng tư) — chỉ ghi chỉ số đoạn, tên lệnh, mã lỗi.
- Thông báo lỗi cho người dùng: 1 câu tiếng Việt + gợi ý; chi tiết kỹ thuật chỉ ở log.

## 6. Bảo mật
- Không `Shell`, `CreateObject("WScript.Shell")`, `MSXML2.XMLHTTP`, `URLDownloadToFile`, ghi registry, tự cập nhật. Cần tính năng cập nhật → ADR + ký số + kênh tin cậy.
- Không tự bật macro/hạ cấp Trust Center. `tools/*.ps1` (chỉ dùng khi phát triển) mới được động tới `AccessVBOM`, và phải **hỏi người dùng trước** (đó là thay đổi cài đặt bảo mật).
- Mọi đầu vào từ văn bản (chuỗi) coi là dữ liệu, không bao giờ `Eval/Application.Run` từ chuỗi đó.

## 7. Giao diện
- Ribbon bằng `customUI14.xml` (nhóm: **Đánh số · Định dạng · Mục lục · Kiểm tra**). Callback mỏng, chỉ gọi hàm trong `features/`.
- Không modal lồng modal. Lệnh phá hủy lớn → hộp xác nhận kèm số đoạn sẽ bị tác động.
- Xem trước (preview) dùng hộp liệt kê, không đổi văn bản cho tới khi bấm OK.

## 8. Hiệu năng
- Duyệt `Paragraphs` bằng `For Each` hoặc lấy `Range` rồi `.Paragraphs`, tránh `Paragraphs(i)` trong vòng lặp lớn (O(n²)).
- Gom thay đổi: áp Style cho cả `Range` thay vì từng đoạn khi có thể.
- Lệnh trên văn bản > 500 đoạn hiện `StatusBar` tiến độ; có thể hủy bằng `DoEvents`+cờ.
- `Scripting.Dictionary` chậm hẳn khi chứa hàng chục nghìn mục trở lên (đo 2026-11-26: thêm 136 nghìn khóa 2,0 s, tra 9 µs/lần so với 1 µs ở dict 1.000 mục). Bảng lớn (bigram) dùng bảng băm địa chỉ mở bằng mảng (`HashS`/`HSlot` trong `modSpell`).
- Hàm gọi theo từng từ/ký tự nhận chuỗi đoạn văn `ByRef` (chỉ đọc), tránh sao chép cả đoạn mỗi lần gọi (O(n²) với đoạn dài).
- Kết quả chỉ phụ thuộc chữ (âm tiết, cặp âm tiết) thì ghi nhớ theo chữ trong một lần chạy; phần phụ thuộc vị trí (`Glued`, chống báo trùng) tính riêng, không ghi nhớ.
- Hằng hex trong VBA: `&HFFFF` là Integer = **-1**; giá trị ≥ &H8000 phải có hậu tố Long `&` (`&HFFFF&`).
- Đo hiệu năng trên laptop dao động 2–3 lần theo tải/nhiệt; giữ mức thường ≤ một nửa ngưỡng KB-E1 để test không chập chờn.

## 9. Checklist tự review trước khi giao
- [ ] `Option Explicit`, không còn biến Variant ngầm.
- [ ] Không `Selection` ghi; không chỉ số đoạn cứng.
- [ ] UndoRecord + nhánh Cleanup khôi phục `ScreenUpdating`, `UndoRecord`.
- [ ] Không số gõ tay; list gắn Style.
- [ ] Chuỗi tiếng Việt qua `modStrings`.
- [ ] Idempotent: chạy 2 lần = 1 lần.
- [ ] Không đụng bảng/header/footer ngoài ý muốn (có test).
- [ ] Không có mạng/Shell.
- [ ] Có kịch bản `KB-xx` và test.
- [ ] Hàm ≤ 40 dòng, tên rõ nghĩa.

## 10. Git & commit
- Commit nhỏ, thông điệp: `KB-A1: áp bộ đánh số Hợp đồng` (tiền tố ID kịch bản).
- Không commit `dist/`, `*.docx` có dữ liệu thật, log.
- Không `--no-verify`, không force push.
