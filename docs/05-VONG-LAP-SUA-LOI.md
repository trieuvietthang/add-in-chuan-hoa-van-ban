# 05 — Vòng lặp tự phát hiện và sửa lỗi

Dùng khi test đỏ, khi Reviewer báo lỗi nặng, hoặc người dùng báo lỗi.

## Vòng lặp (tối đa 3 vòng cho mỗi lỗi)
```
1. TÁI HIỆN   → có fixture/lệnh làm lỗi xuất hiện ổn định (nếu không tái hiện được: dừng, hỏi dữ liệu)
2. KHOANH VÙNG→ dùng tools/inspect.ps1 so trước/sau; xác định đoạn/hàm đầu tiên cho kết quả sai
3. GIẢ THUYẾT → viết 1–2 giả thuyết nguyên nhân gốc, chọn cái dễ kiểm chứng nhất
4. KIỂM CHỨNG → thêm log/assert hoặc thử nghiệm nhỏ; xác nhận giả thuyết
5. SỬA        → sửa ở nguyên nhân gốc, thay đổi nhỏ nhất
6. HỒI QUY    → thêm test cho lỗi này; chạy TOÀN BỘ test
7. KẾT THÚC   → xanh hết → ghi vào "Nhật ký lỗi" (dưới); còn đỏ → quay lại bước 2 (vòng +1)
```
Quá 3 vòng, hoặc sửa làm vỡ test khác 2 lần liên tiếp → **dừng**, báo người dùng: triệu chứng, đã thử gì, giả thuyết còn lại, đề xuất.

## Quy tắc
- Không xóa/nới test để "xanh". Test sai thì chứng minh sai (kịch bản đổi) rồi sửa kèm lý do.
- Không bọc `On Error Resume Next` để nuốt lỗi.
- Không sửa nhiều thứ cùng lúc; mỗi vòng một thay đổi.
- Sửa xong phải build lại `dist/` và chạy lại từ đầu (tránh test trên bản cũ).
- Mọi dữ liệu thật dùng để tái hiện phải ẩn danh trước khi đưa vào `tests/fixtures`.

## Bảng chẩn đoán nhanh (Word/VBA hay gặp)
| Triệu chứng | Nguyên nhân thường gặp | Hướng xử lý |
|---|---|---|
| Số không tự cập nhật khi chèn Điều mới | Số gõ tay, hoặc đoạn mới không gắn Style của list | Gắn LinkedStyle; chạy KB-A4 |
| Số nhảy 1,2,1,2 | `ResetOnHigher` sai hoặc dùng 2 ListTemplate khác nhau | Dùng 1 ListTemplate đa cấp; kiểm tra `ListLevels(i).ResetOnHigher` |
| Thụt dòng lệch giữa số và chữ | `NumberPosition/TextPosition/TabPosition` không khớp; ký tự theo sau không phải Tab | Đặt cả 3 + `TrailingCharacter = wdTrailingTab` |
| Ctrl+Z phải bấm nhiều lần | Thiếu `UndoRecord` hoặc `EndCustomRecord` không chạy khi lỗi | Dùng mẫu ở `02-LUAT-CODE` mục 3 |
| Word đứng/chậm | `Paragraphs(i)` trong vòng lặp lớn; `ScreenUpdating` bật | Duyệt `For Each`; tắt ScreenUpdating |
| Chữ tiếng Việt thành `?`/lạc | Chuỗi dấu hard-code trong file ANSI | Dùng `modStrings` |
| Style "Heading 1" không tìm thấy trên Word tiếng Việt | Gọi bằng tên | Dùng `wdStyleHeading1` hoặc style `VB-` |
| TOC trống | Đoạn không có outline level / chưa dùng Heading/`VB-` có `OutlineLevel` | Gán `ParagraphFormat.OutlineLevel`; `TablesOfContents.Add UseHeadingStyles/AddedStyles` |
| Mất định dạng ký tự | `Range.Font.Reset` hoặc `ClearFormatting` quá tay | Chỉ chạm định dạng đoạn |
| Lệnh lỗi trong bảng/vùng bảo vệ | Giả định `Range` liền khối | Kiểm tra `Range.Information(wdWithInTable)`, `ProtectionType` |
| Test bị treo vì `WINWORD.EXE` mồ côi | Script không đóng Word khi lỗi | `try/finally` + `Get-Process WINWORD` dọn dẹp (chỉ tiến trình do script mở) |
| COM từ chối `VBProject` | `AccessVBOM` tắt | Hỏi người dùng trước khi bật; ghi cách tắt lại |

## Nhật ký lỗi (điền khi sửa xong)
| Ngày | ID lỗi | Triệu chứng | Nguyên nhân gốc | Cách sửa | Test hồi quy |
|---|---|---|---|---|---|
| | | | | | |

## Báo cáo khi gặp lỗi (mẫu gửi người dùng)
```
Lỗi: <1 câu>
Tái hiện: <fixture + lệnh>
Nguyên nhân gốc: <đã xác nhận / giả thuyết>
Đã sửa: <file:hàm> — Test hồi quy: <tên>
Kết quả test: <x PASS / y FAIL>
Chưa kiểm chứng: <nếu có>
```
