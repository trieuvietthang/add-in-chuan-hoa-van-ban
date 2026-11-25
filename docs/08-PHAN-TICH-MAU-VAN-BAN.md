# 08 — Phân tích cấu trúc 3 văn bản mẫu của người dùng

> Phương pháp: đọc **chỉ cấu trúc** (style, định nghĩa đánh số, trang, field, định dạng trực tiếp) bằng script đọc-only trên OOXML; **không đọc/ghi nhớ nội dung**, không sửa file gốc. Không trích nội dung vào tài liệu này. Ngày phân tích: 2026-11-23.
> Ký hiệu: **F1** = hợp đồng cung cấp LMS (≈2,3 MB) · **F2** = báo cáo pháp lý · **F3** = quy chế tổ chức lấy ý kiến điện tử.
> Đơn vị: twip (1 cm ≈ 567 twip; 1 pt = 20 twip); `sz` = nửa điểm (24 → 12 pt).

## 1. Bảng so sánh nhanh
| Hạng mục | F1 Hợp đồng | F2 Báo cáo pháp lý | F3 Quy chế |
|---|---|---|---|
| Khổ giấy | A4 | **Letter** (12240×15840) | A4 |
| Lề (trên/dưới/trái/phải) | 2,54 cm cả 4 phía | 2,54 cm cả 4 phía | 2,0 / 1,8 / **2,5** / 2,0 cm |
| Số section | **6** (cùng khổ, cùng lề → thừa) | 1 | 1 (+ 3 header/footer) |
| Font mặc định tài liệu | **Inter** (docDefaults) | Calibri 11 | không khai báo |
| Font thực tế của style | TNR 12 (đặt riêng ở style điều khoản) | Calibri 11 (Normal); Heading 16/13/12 đậm | TNR 12 |
| Số đoạn / bảng | 404 / 8 | 217 / 6 | 118 / 1 |
| Style kiểu "Chương–Điều" | `01Dieu`, `02Khoản`, `03Diem`, `04Tiet` | không dùng (dùng Heading 1–3) | `00Chuong`, `01Dieu`, `02Khoan`, `03Diem` |
| Đoạn dùng style điều khoản | 147/404 (Điều 14, Khoản 57, Điểm 66, Tiết 10) | 0 | 59/118 |
| Đoạn "Normal" | 228 | 180 | 16 |
| Mục lục | **Không** | **Không** | Có: `TOC \t "Heading 1,1,Heading 2,2" \h \z` |
| Tham chiếu chéo (REF) | 0 (20 bookmark) | 0 | 0 (22 bookmark do TOC) |
| Footnote / comment | có / **có** | có / không | có / không |
| Số trang | footer `PAGE`/`NUMPAGES` | không thấy | footer `PAGE`+`NUMPAGES` trong bảng 1 hàng |
| Đoạn có `numPr` gắn trực tiếp (bỏ qua style) | 36 | 31 | 1 |
| Định dạng đoạn trực tiếp (ind/spacing/jc) | 64 / 257 / 257 | 52 / 157 / 123 | 14 / 18 / 20 |
| Định dạng ký tự trực tiếp (rFonts / sz / color) | **715 / 714 / 713** | **868 / 829 / 9** | 125 / 91 / 3 |
| Số `abstractNum` / `num` | 19 / 19 | 15 / 19 | **23 / 23** |

## 2. Cách người dùng đang làm (suy ra từ cấu trúc)
1. **Một danh sách đa cấp duy nhất gắn với style** (đúng kỹ thuật chuẩn). Ví dụ F3: một `abstractNum` đa cấp có 6 cấp, các cấp 0–3 gắn `LinkedStyle` = `00Chuong` → `01Dieu` → `02Khoan` → `03Diem`:
   - F3: `CHƯƠNG %1` (La Mã) · `Điều %2.` · `%3.` · `%4)` · `(%5)` · `(%6)`
   - F1: `Điều %1.` · `%1.%2` (kiểu 1.1) · `%3)` (a) · `(%4)` (i, La Mã thường)
   - F2 (báo cáo): `I.` · `a.` · `(1)` … nhưng làm bằng danh sách gắn trực tiếp, **không** qua style.
2. **Đặt tên style có tiền tố số** (`00…04`) để sắp thứ tự trong bảng Styles; style kế thừa từ Heading 1/2/3 (F3) hoặc Normal (F1).
3. **Style chỉ mô tả thụt dòng + font + đậm**: Điều đậm, thụt treo (hanging) 851 twip ≈ 1,5 cm (F3) hoặc không; Khoản treo 567 (≈1 cm) ở F3 / treo 720 ở F1; Điểm thụt trái 567/720.
4. Văn bản **không đồng nhất giữa các file**: kiểu số Khoản (`1.` ở F3 nhưng `1.1` ở F1), cấp của Điều (cấp 0 ở F1, cấp 1 dưới Chương ở F3), font (TNR/Inter/Calibri), lề, khổ giấy.
5. Mục lục: chỉ F3 có; dùng Heading 1/2 làm nguồn, hai style `toc 1`, `toc 2` chỉnh tay.

## 3. Vấn đề phát hiện (đây chính là việc add-in phải tự động hóa)
| # | Vấn đề | Bằng chứng | Hệ quả | Cách add-in xử lý |
|---|---|---|---|---|
| P1 | **Lẫn định dạng trực tiếp với style** | F1: 715 run có `rFonts`, 713 có `color`; 257 đoạn chỉnh giãn/căn trực tiếp | Dán từ nguồn khác mang font/màu lạ; đổi style không có tác dụng | Chuẩn hóa: đặt Normal/style đúng, xóa định dạng trực tiếp thừa (giữ đậm/nghiêng/gạch chân có chủ đích) |
| P2 | **Font mặc định sai** (Inter ở F1, Calibri ở F2) | docDefaults | Đoạn không gán style tự rơi về Inter/Calibri | Đặt font ở style Normal + các style của hồ sơ; không phụ thuộc docDefaults |
| P3 | **Đánh số bỏ qua style** | F1 36 đoạn, F2 31 đoạn có `numPr` trực tiếp | Không tăng/giảm cấp bằng style; không vào mục lục; số lệch khi chèn | Chuyển về style + list template của hồ sơ |
| P4 | **Danh sách "rác"** | 19–23 `abstractNum` mỗi file | Tăng dung lượng; dễ nhảy số/nhầm list khi dán | Dùng lại đúng 1 list template; báo số list mồ côi; (tùy chọn) dọn bằng thao tác riêng có xác nhận |
| P5 | **Khổ giấy/lề không chuẩn** | F2 là Letter; F3 lề trái 2,5 cm; F1 lề 2,54 | In ra sai khổ, không đúng NĐ 30/ISO | Hồ sơ trang: A4 + lề theo chuẩn chọn |
| P6 | **Section thừa** | F1 có 6 section giống hệt nhau | Header/footer rối, số trang lệch | Phát hiện section trùng thông số, đề xuất gộp (không tự gộp nếu có header/footer khác nhau) |
| P7 | **Khoản gắn vào Heading** | F3: `02Khoan` kế thừa Heading 3 | Đoạn dài hiện như "heading" trong Navigation Pane; có thể lọt TOC | Style điều khoản đặt `OutlineLevel` riêng: Chương/Mục/Điều có cấp, Khoản/Điểm/Tiết = văn bản thường |
| P8 | **TOC chọn theo tên Heading 1/2** (F3) nhưng đoạn thực là style tùy biến | `\t "Heading 1,1,Heading 2,2"` | TOC có thể thiếu/dư mục tùy kế thừa; phải kiểm chứng | TOC lập theo `\t` liệt kê style của hồ sơ + `\u` (outline) |
| P9 | **Không có tham chiếu chéo thật** | 0 trường REF | Số "Điều X" gõ tay → sai khi đánh số lại (rủi ro pháp lý) | Kiểm tra dẫn chiếu + chuyển thành REF tự cập nhật |
| P10 | **Không thống nhất giữa loại văn bản** | Khoản `1.` vs `1.1`, số cấp khác nhau | Khó dùng một thao tác | Mỗi loại văn bản = một **hồ sơ (profile)** có bộ style/số riêng |
| P11 | **Trang đầu/ số trang** | Footer chứa bảng; F2 không số trang | Khó tự động | Hồ sơ cho phép kiểu số trang: "Trang X/Y", "từ trang 2 ở giữa đầu trang" |
| P12 | **Comment + footnote có trong tài liệu** | F1 có comment; cả 3 có footnote | Phải bảo toàn khi chuẩn hóa | Test bắt buộc: không mất comment/footnote |

## 4. Thao tác thủ công hiện tại → tự động hóa
| Thao tác bạn làm tay | Add-in thay bằng |
|---|---|
| Tạo style Chương/Điều/Khoản/Điểm/Tiết + Multilevel List, liên kết từng cấp | **Cài hồ sơ**: tạo/sửa style và list template một lần, đúng thông số |
| Chọn từng đoạn, bấm style tương ứng | **Nhận diện cấu trúc** + xem trước + áp style hàng loạt; phím tắt/nút "Chương/Điều/Khoản/Điểm/Tiết" áp cho vùng chọn; Tăng/Giảm cấp |
| Gõ tay `I.`, `1.`, `1.1`, `a)`, `(i)`, `-` cho công văn | Hồ sơ **Công văn** có list nhiều cấp đúng ký hiệu |
| Sửa font/cỡ/giãn/lề thủ công | **Chuẩn hóa trang & phông** theo hồ sơ (NĐ 30 / ISO 9001 / tùy chỉnh) |
| Lập mục lục, chỉnh toc style | **Mục lục** theo hồ sơ + cập nhật |
| Rà lỗi chính tả, dẫn chiếu "Điều X" bằng mắt | **Kiểm tra**: chính tả (luật + từ điển âm tiết), dẫn chiếu, cấu trúc, thuật ngữ xác định |

## 5. Việc không rõ cần xác nhận với người dùng (đã đưa vào `06-NGHIEP-VU` mục câu hỏi mở)
- Chuẩn kiểu số Khoản cho **hợp đồng**: `1.1` (như F1) hay `1.` (như F3)? Có cho phép cấu hình theo từng hồ sơ không (mặc định: có).
- Công văn: dãy ký hiệu cấp bạn mong muốn chính xác là `I. → 1. → 1.1. → a) → (i) → -`? Cấp "Mục/Phần" có thuộc công văn không?
- Lề chuẩn ISO 9001 của đơn vị là bao nhiêu (NĐ 30 là 2,0–2,5 trên/dưới; 3,0–3,5 trái; 1,5–2,0 phải)?
- Mục lục hợp đồng: chỉ Điều, hay cả Phụ lục? Có số trang không?
