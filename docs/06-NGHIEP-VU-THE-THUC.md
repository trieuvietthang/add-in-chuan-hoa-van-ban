# 06 — Nghiệp vụ thể thức & đánh số (Việt Nam)

> **Lưu ý:** các thông số dưới đây là giá trị mặc định tham khảo, tổng hợp từ hiểu biết về NĐ 30/2020/NĐ-CP (công tác văn thư). **Chưa đối chiếu lại với văn bản gốc trong phiên này.** Trước khi hard-code, Analyst phải đối chiếu văn bản hiện hành và cập nhật cột "Đã xác minh". Mọi giá trị phải nằm trong **bảng cấu hình** (không rải rác trong code) để đổi theo quy chế từng cơ quan.

## 1. Công văn / văn bản hành chính (NĐ 30/2020)
| Thông số | Mặc định | Đã xác minh |
|---|---|---|
| Khổ giấy | A4 (210×297 mm), dọc | ☐ |
| Lề trên / dưới | 20–25 mm (mặc định 20) | ☐ |
| Lề trái | 30–35 mm (mặc định 30) | ☐ |
| Lề phải | 15–20 mm (mặc định 20 hoặc 15) | ☐ |
| Font | Times New Roman (Unicode), màu đen | ☐ |
| Cỡ chữ nội dung | 13–14 pt (mặc định 14) | ☐ |
| Thụt đầu dòng đoạn văn | 1 cm – 1,27 cm (mặc định 1,27) | ☐ |
| Giãn dòng | từ đơn (single) đến 1,5 dòng (mặc định ~1,15–1,5) | ☐ |
| Giãn đoạn | trước/sau ≥ 6 pt (mặc định 6) | ☐ |
| Căn lề đoạn | Justified (đều hai bên) | ☐ |
| Số trang | Ả Rập, cỡ 13–14, giữa đầu trang (header), **ẩn trang 1** | ☐ |
| Quốc hiệu/tiêu ngữ | khối đầu trang, theo mẫu thể thức | ☐ |

## 2. Hợp đồng (không có quy chuẩn pháp luật bắt buộc về cách đánh số → theo thông lệ)
Mặc định `profile = HopDong`, cho phép cấu hình:
| Cấp | Style | Định dạng số | Ví dụ |
|---|---|---|---|
| 1 | `VB-Dieu` | `Điều %1.` (Ả Rập) + tiêu đề điều in đậm | `Điều 1. Đối tượng hợp đồng` |
| 2 | `VB-Khoan` | `%1.%2.` | `1.1. Bên A có trách nhiệm…` |
| 3 | `VB-Diem` | `%3)` chữ thường a, b, c | `a) …` |
| 4 | `VB-Gach` | gạch đầu dòng `-` | `- …` |
- Phụ lục: đánh số riêng (ListTemplate riêng hoặc `Restart`), tiêu đề `PHỤ LỤC 01`.
- Phương án B (tùy chọn): khoản chỉ `1.`, `2.` (không kèm số Điều).

## 3. Quy định / quy chế nội bộ
Mặc định `profile = QuyDinh`:
| Cấp | Style | Định dạng số | Ví dụ |
|---|---|---|---|
| 1 | `VB-Chuong` | `Chương %1` (La Mã), tên chương xuống dòng, in đậm, căn giữa | `Chương I` / `QUY ĐỊNH CHUNG` |
| 2 | `VB-Muc` (tùy chọn) | `Mục %2.` | `Mục 1.` |
| 3 | `VB-Dieu` | `Điều %3.` **đếm liên tục toàn văn bản** | `Điều 5. Phạm vi áp dụng` |
| 4 | `VB-Khoan` | `%4.` | `1. …` |
| 5 | `VB-Diem` | `%5)` a, b, c | `a) …` |
| 6 | `VB-Gach` | `-` | |
- Quyết định kỹ thuật: số Điều liên tục xuyên Chương ⇒ `ResetOnHigher` của cấp Điều = **tắt** (0) so với Chương/Mục; cấp Khoản reset theo Điều.
- Tên chương 2 dòng ("Chương I" + dòng tên) cần xử lý: số ở dòng 1 bằng list, dòng tên là đoạn style riêng `VB-TenChuong` (outline level cùng cấp để TOC gộp được) — ghi ADR khi chốt.

## 4. Quy tắc định dạng chung
- Số liệu thụt: `NumberPosition` (vị trí số) = thụt của cấp cha; `TextPosition` = vị trí chữ sau số; khoảng cách số–chữ bằng 1 Tab (không dùng dấu cách).
- Gợi ý thụt (cm): Điều 0 / 1,27; Khoản 0 / 1,27 hoặc 1,27 / 2,0; Điểm 1,27 / 2,0; Gạch 2,0 / 2,5 (cấu hình được).
- Tiêu đề điều in đậm là **định dạng ký tự trong Style**, không bôi tay.
- Dấu: ngoặc kép kiểu thẳng/cong theo lựa chọn; không dấu cách trước `, ; : . ! ?`; một dấu cách sau dấu câu.
- Số/Ngày: `ngày 05 tháng 11 năm 2026`; không viết tắt "ngày …/…/…" trong bản chuẩn trừ khi chọn.

## 5. Mục lục
- Hợp đồng: cấp 1 = Điều (`TOC \o`/`\t "VB-Dieu,1"`), tùy chọn không số trang.
- Quy định: cấp 1 = Chương (+ tên chương), cấp 2 = Mục, cấp 3 = Điều; tuỳ chọn ẩn cấp Điều.
- Dùng trường `TOC \h \z \t "VB-Chuong,1,VB-Muc,2,VB-Dieu,3"` (tên style liệt kê phải khớp danh sách tách bằng **dấu phẩy hoặc chấm phẩy theo cài đặt vùng** của Windows — kiểm tra bằng `Application.International(wdListSeparator)`).

## 7. Quy ước thực tế của người dùng (suy từ 3 mẫu — xem `08`)
Các giá trị dưới đây là **mặc định ban đầu của Hồ sơ**, lấy từ cách bạn đang làm; bạn chỉnh được.
| Hồ sơ | Cấp (style) | Định dạng số | Ghi chú |
|---|---|---|---|
| `HopDong` | `01Dieu` / `02Khoan` / `03Diem` / `04Tiet` | `Điều %1.` / `%1.%2` / `%3)` (a) / `(%4)` (i) | Giống F1; Điều đậm, justify, TNR 12 |
| `QuyDinh` | `00Chuong` / (`00Muc`) / `01Dieu` / `02Khoan` / `03Diem` / `04Tiet` | `CHƯƠNG %1` La Mã / `Mục %n` / `Điều %n.` / `%n.` / `%n)` / `(%n)` | Giống F3; Chương căn giữa đậm; Điều đếm liên tục; thụt treo Điều ≈ 1,5 cm, Khoản/Điểm ≈ 1 cm |
| `CongVan` | `CV1`…`CV6` (đã chốt ở M1) | `I.` / `1.` / `1.1.` / `a)` / `(i)` / `-` | Theo mô tả của bạn; xác nhận ở câu hỏi 8 |
- Chuẩn trang chọn được: **NĐ 30** (mục 1) hoặc **ISO 9001** (theo quy định đơn vị, điền ở câu 3) — mặc định chung: A4, Times New Roman 12 hoặc 13, giãn dòng **Multiple 1.15 hoặc 1.2**, giãn đoạn sau 6 pt, căn đều hai bên.
- Style điều khoản có **`OutlineLevel` riêng**: Chương 1, Mục 2, Điều 3 (đưa vào mục lục); Khoản/Điểm/Tiết = văn bản thường (không hiện ở Navigation/TOC), để tránh lỗi P7.
- Nhận diện tên style cũ (`legacyAliases`): `01Dieu`, `02Khoản`/`02Khoan`, `03Diem`, `04Tiet`, `00Chuong`, `Heading 1–3`.

## 6. Câu hỏi mở cho người dùng (điền khi trả lời)
1. Công ty đang dùng kiểu đánh số hợp đồng nào (Điều 1 / 1.1 / a hay khác)? ______
2. Quy định nội bộ có dùng "Mục" không? Số Điều liên tục xuyên Chương? ______
3. Lề, font, cỡ chữ chuẩn nội bộ (nếu khác NĐ 30)? ______
4. Phụ lục hợp đồng đánh số thế nào? ______
5. Có dùng song ngữ Việt–Anh không? ______
6. Phiên bản Word (2016/2019/365) và có bị chặn macro không? ______
7. Số Khoản trong **hợp đồng**: `1.1` (như mẫu hợp đồng) hay `1.` (như mẫu quy chế)? Cho phép cấu hình theo hồ sơ (mặc định: có). ______
8. Công văn: dãy ký hiệu chính xác `I. → 1. → 1.1. → a) → (i) → -`? Có cấp "Mục/Phần"? ______
9. Lề ISO 9001 của đơn vị (trên/dưới/trái/phải, cm)? Cỡ chữ 12 hay 13 làm mặc định? Giãn dòng 1.15 hay 1.2? ______
10. Mục lục hợp đồng: chỉ Điều, hay cả Phụ lục? Có số trang/dẫn chấm? ______
11. Đồng ý bật `AccessVBOM` (Trust access to VBA project) trên máy dev để Claude tự build/test (tắt lại được)? ______

## 8. Cấu hình mặc định đã triển khai (M1)
Xem config/profiles/*.ini. Quyết định tạm (đổi trong INI): Điều/Chương/Mục/Công văn dùng dấu cách sau số (Điều 1. Tên, CHƯƠNG I Tên); Khoản/Điểm/Tiết dùng Tab; HĐ Khoản 1.1; Quy định Khoản 1.; Chương CHƯƠNG I La Mã. Nhận diện số gõ tay bằng biểu thức chính quy detect= trong từng cấp.
