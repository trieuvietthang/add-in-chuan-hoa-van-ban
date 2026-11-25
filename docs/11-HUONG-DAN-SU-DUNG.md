# Hướng dẫn sử dụng add-in "Soạn thảo chuẩn"

Add-in cho Word (Windows) giúp đánh số điều khoản, chuẩn hóa thể thức, làm mục lục, kiểm tra dẫn chiếu và chính tả cho **Hợp đồng · Quy định nội bộ · Công văn**. Chạy hoàn toàn offline, không gửi dữ liệu ra ngoài.

## 1. Cài đặt
1. Giải nén gói `SoanThaoChuan-<phiên bản>.zip`, **đóng Word**.
2. Chạy (PowerShell):
   `powershell -ExecutionPolicy Bypass -File install.ps1`
   Script nói rõ sẽ ghi gì rồi hỏi xác nhận. Chỉ ghi `SoanThaoChuan.dotm` và thư mục `config` vào `%APPDATA%\Microsoft\Word\STARTUP`. Cấu hình cũ được sao lưu ở `%APPDATA%\SoanThaoChuan\backup\`.
3. Mở Word → thẻ **Soạn thảo chuẩn**.
4. Gỡ: `install.ps1 -Uninstall`. Chạy thử không ghi gì: thêm `-DryRun`.

**Nếu Word chặn macro** (thanh vàng/đỏ "Macros have been disabled" hoặc thẻ không hiện): đây là chính sách bảo mật của Word/công ty, không phải lỗi add-in. Cách xử lý: nhờ IT ký số tệp `.dotm` bằng chứng chỉ của công ty (hướng dẫn và mẫu gửi IT ở `docs/12-KY-SO.md`), hoặc thêm thư mục STARTUP vào *Trusted Locations* (File → Options → Trust Center). Add-in không tự đổi cài đặt bảo mật. `install.ps1` cho biết tệp đã có chữ ký VBA hay chưa.

## 2. Quy tắc an toàn
- Add-in **không sửa nội dung chữ** của bạn (trừ khi bạn bấm lệnh chuyển/chuẩn hóa và đã xem thông báo). Mỗi lệnh ghi vào văn bản có thể hoàn tác bằng **một lần Ctrl+Z**.
- Nên làm việc trên **bản sao** khi thử lần đầu với văn bản quan trọng.

## 3. Các nhóm lệnh
| Nhóm | Lệnh | Việc làm |
|---|---|---|
| Hồ sơ | Loại văn bản | Chọn Hợp đồng / Quy định nội bộ / Công văn. Quyết định kiểu đánh số và style. |
| Đánh số | Nút cấp 1…6, Tăng/Giảm cấp | Gán cấp (Chương/Mục/Điều/Khoản/Điểm/Tiết, hoặc I. 1. 1.1. a) (i) với công văn) cho đoạn đang chọn. Số tự cập nhật khi chèn/xóa. |
| Chuyển đổi | Số gõ tay → tự động | Tìm số gõ tay ("Điều 1.", "1.1", "a)"), cho xem danh sách, rồi chuyển thành đánh số thật. |
| Mục lục | Chèn / Cập nhật | Mục lục chuẩn theo các cấp của hồ sơ; số trang tự cập nhật. |
| Dẫn chiếu | Kiểm tra, Lỗi tiếp theo, Chuyển thành tự động, Cập nhật, Chèn dẫn chiếu Điều | Phát hiện dẫn chiếu "Điều N" trỏ tới Điều không có, trùng tên Điều khác, khoản/điểm không tồn tại; chuyển dẫn chiếu gõ tay thành dẫn chiếu tự động đổi số theo Điều. Dẫn chiếu tới văn bản khác ("Điều 62 Quyết định số …") được bỏ qua. |
| Chính tả | Kiểm tra chính tả, Lỗi tiếp theo | Chỉ **báo**, không tự sửa. Xem mục 4. |
| Chuẩn hóa | Chuẩn (NĐ 30/2020 hoặc ISO 9001), Chuẩn hóa trang & phông | Lề, khổ giấy, phông, cỡ chữ, giãn dòng, thụt dòng, số trang. Chữ không đổi. Có thông báo tóm tắt trước khi áp dụng. |

## 4. Kiểm tra chính tả
Mã lỗi hiện trong hộp thoại tóm tắt:
- **S1–S6 (kiểu chữ):** hai dấu cách, dấu cách trước dấu câu, thiếu dấu cách sau dấu câu, lặp từ liền kề, ngoặc/nháy không cân, "điều" viết thường đầu câu.
- **W1:** cặp sai–đúng thường gặp ("xem sét"→"xem xét").
- **W2:** cùng một từ viết hai cách ("quy"/"qui" định); báo cách ít dùng hơn.
- **W3:** âm tiết không hợp lệ trong tiếng Việt (gõ sai chữ/dấu, như "hợpp").
- **W4:** thuật ngữ định nghĩa (Bên A/B/C) sai hoa/thường. Chữ IN HOA toàn bộ được bỏ qua.
- **W5:** cặp âm tiết lạ nghi gõ sai dấu ("quy đinh" → gợi ý "quy định").

**Giới hạn cần biết:** không hiểu nghĩa câu; không bắt lỗi đúng âm tiết nhưng sai từ (hỏi/ngã) khi cặp từ vẫn phổ biến; không quét chú thích cuối trang và đầu/cuối trang; không đối chiếu số tiền với số bằng chữ. Hộp thoại không hiện được dấu tiếng Việt (giới hạn của Word), nên gợi ý hiện không dấu — xem chữ đúng tại vị trí lỗi.

## 5a. Sửa hồ sơ bằng form
Thẻ **Soạn thảo chuẩn** → nhóm Hồ sơ → **Sửa hồ sơ…**. Chọn *Tệp* (Chuẩn NĐ 30, Chuẩn ISO 9001, từng Hồ sơ, hoặc Chính tả) rồi chọn *Mục*; sửa các ô và bấm **Lưu**.
- Giá trị được kiểm tra trước khi ghi (số trong khoảng cho phép, kiểu số/căn lề thuộc danh sách, in đậm chỉ 0/1…). Giá trị sai thì không ghi và form báo ô nào sai.
- Tệp gốc được sao lưu thành `.bak` cạnh tệp; chú thích và thứ tự trong tệp được giữ nguyên.
- Một số ô chỉ đọc (mã hồ sơ, số cấp, tên style, loại dẫn chiếu) vì đổi sẽ làm lệch văn bản đã đánh số.
- Giá trị mới có hiệu lực ở lần áp dụng kế tiếp: sau khi đổi, chạy lại lệnh (Chọn cấp/Chuẩn hóa/Mục lục).
- Mục chính tả trong form chỉ gồm ngưỡng W5, từ bỏ qua, thuật ngữ, từ láy hợp lệ. Cặp sai–đúng và nhóm biến thể vẫn sửa bằng Notepad (mục 5).

## 5. Tùy biến bằng tệp (không cần build lại)
Thư mục `config` cạnh `SoanThaoChuan.dotm` (trong STARTUP). Sửa bằng Notepad, lưu UTF-8, mở lại Word.
- `config\spell.vi.ini` — `[pairs]` thêm cặp sai=đúng; `[terms]` thuật ngữ; `[groups]` nhóm biến thể; `[ignore] words=` từ/cặp bỏ qua (ví dụ tên riêng, viết tắt; cặp ghi `am1 am2`, cách nhau dấu phẩy); `[w5]` ngưỡng nhạy (`min_alt`, `min_ratio`; tăng để ít báo hơn; `enabled=0` để tắt W5).
- `config\profiles\*.ini` — kiểu đánh số/tên style từng hồ sơ; `config\standards\*.ini` — số liệu NĐ 30 / ISO 9001 (lề, phông, giãn dòng).
- `config\refs.vi.ini` — mẫu nhận diện dẫn chiếu.
- `config\bigram.vi.txt` — dữ liệu W5 (sinh bằng `tools\build-bigrams.ps1`, nguồn ghi ở mục 7). Không cần sửa tay.

**Xuất/nhập cấu hình:** sao chép nguyên thư mục `config` sang máy khác (hoặc dùng gói cài đặt). Mỗi lần cài, bản cũ tự được sao lưu.

## 6. Xử lý sự cố
| Hiện tượng | Cách xử lý |
|---|---|
| Không thấy thẻ "Soạn thảo chuẩn" | Macro bị chặn (mục 1), hoặc Word chưa mở lại sau khi cài. |
| "Không đọc được tệp cấu hình" | Thiếu thư mục `config` cạnh `.dotm`; cài lại bằng `install.ps1`. |
| Chính tả báo nhầm từ đúng (tên riêng, viết tắt, thuật ngữ pháp lý hiếm) | Thêm vào `[ignore]`; hoặc tăng `min_alt`/`min_ratio` nếu báo nhầm W5 nhiều. |
| Mục lục trống | Chưa có đoạn nào gán cấp; dùng "Số gõ tay → tự động" hoặc các nút cấp trước. |
| Lệnh báo lỗi giữa chừng | Văn bản được giữ nguyên trạng thái trước; nếu có thay đổi dở, bấm Ctrl+Z. |

## 7. Nguồn dữ liệu và giấy phép
- Bảng `bigram.vi.txt` thống kê từ phần câu đúng của bộ ngữ liệu `bmd1905/error-correction-vi` (Hugging Face), giấy phép Apache-2.0. Add-in chỉ chứa bảng tần suất âm tiết, không chứa câu gốc.
- Mã nguồn add-in viết mới cho dự án này.
