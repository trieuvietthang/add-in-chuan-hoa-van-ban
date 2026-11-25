# 00 — Đánh giá mức độ tương thích của repo `CuongBC195/add-in-office`

> Ngày đánh giá: 2026-11-23. Nguồn: README repo + GitHub Releases API + `version.json` bản v2026.11.20.3.
> **Chưa** tải/mở/giải nén file `.dotm`/`.xlam` nào (xem mục 5).

## 1. Repo thực sự chứa gì

| Mục | Kết quả |
|---|---|
| Nhánh `main` | Chỉ có **1 file `README.md`**. Không có mã nguồn, không có license. |
| Mục đích (theo README) | "Repo này chỉ dùng để phát hành bản cài đặt và bản cập nhật" |
| Sản phẩm | Add-in **"Chuẩn hóa Thể thức Văn bản"** cho Word + căn chỉnh trang in A4 cho Excel |
| Hình thức | `ChuanHoa_VanBanDang.dotm` (Word, template có macro) và `ChuanHoa_Excel.xlam` (Excel), có biến thể `_Win` và bản Mac |
| Phân phối | Release asset: `Bo_Cai_Dat_AddIn_Windows.zip`, `Bo_Cai_Dat_AddIn_Mac.zip` (~240 KB); nút **Cập nhật** trên Ribbon đọc `version.json` (có `sha256`, `size`, URL) |
| Nhịp phát hành | ~20 release từ 05/11 → 20/11/2026 (gần như mỗi ngày); kích thước `.dotm` tăng 148 KB → 181 KB (đang thêm tính năng đều) |
| Công nghệ suy ra | **VBA trong `.dotm`**, KHÔNG phải Office.js/web add-in (không có manifest, không có host web). Đây là suy luận từ phần mở rộng file — chưa xác minh bên trong. |

**Hệ quả quan trọng:** không thể "fork và sửa tiếp" repo này vì không có source. Có 3 hướng:
1. **A. Viết lại từ đầu** (khuyến nghị) — lấy repo làm tham chiếu chức năng/UX; source là file text `.bas/.cls/.frm` trong repo của bạn.
2. **B. Xin tác giả** (CuongBC195) cung cấp source/cho phép phát triển tiếp — kiểm tra điều khoản sử dụng.
3. **C. Mở `.dotm`, export module VBA** để học — chỉ làm khi đã xác nhận được phép (xem mục 5), không copy nguyên si khi chưa có license.

## 2. Đối chiếu với nhu cầu của bạn

Nhu cầu: sửa/định dạng **Hợp đồng, công văn, quy định nội bộ** — đánh số điều khoản nhanh, format chuẩn, thụt dòng, số thứ tự đúng chuẩn, xuất mục lục.

| Nhu cầu | Repo có phủ không? | Mức tin cậy | Ghi chú |
|---|---|---|---|
| Format **công văn** theo thể thức (NĐ 30/2020/NĐ-CP) | **Có khả năng cao** — tên "Chuẩn hóa Thể thức Văn bản", "VanBanDang" | Trung bình (suy từ tên) | Là hướng chính của repo |
| Căn lề, font, giãn dòng, thụt đầu dòng | Nhiều khả năng có | Trung bình | Chưa kiểm chứng giá trị mặc định |
| **Hợp đồng** (Điều/Khoản/Điểm) | **Không rõ** | Thấp | Release notes không nhắc tới |
| **Quy định nội bộ** (Chương/Mục/Điều) | **Không rõ** | Thấp | Cấu trúc Chương–Điều khác công văn |
| Đánh số **multilevel list gắn style** (đúng chuẩn Word, tự cập nhật khi chèn/xóa) | **Không rõ** | Thấp | Đây là điểm khó nhất, quyết định chất lượng |
| **Mục lục** (TOC field) | **Không rõ** | Thấp | Không có trong mô tả |
| Tham chiếu chéo "theo Điều 5 Khoản 2" tự cập nhật | Không rõ | Thấp | Cần REF field |
| Excel A4 | Có (nhưng ngoài phạm vi của bạn) | Cao | Có thể bỏ qua |
| Mac + Windows | Có (2 bộ cài) | Cao | Bạn đang dùng Windows 11 |

**Kết luận:** mức tương thích **một phần (~30–40% dự đoán)**. Trùng ở lớp "chuẩn hóa thể thức công văn/trang in"; **thiếu hoặc không xác minh được** ở lớp bạn cần nhất: đánh số điều khoản đa cấp, hợp đồng/quy định nội bộ, mục lục. Hơn nữa, vì không có source, repo không thể là nền để mở rộng. Dùng nó như **tham chiếu UX + tham chiếu test** (cài thử, so sánh đầu ra), còn code của dự án này nên tự xây theo `CLAUDE.md`.

## 3. Rủi ro khi dùng/lệ thuộc repo
- **Không có source + không license** → không có quyền sửa/phân phối lại.
- **Macro tự cập nhật từ internet** (tải `.dotm` mới từ GitHub, kiểm tra sha256 nhưng sha256 lấy từ cùng nguồn → không chống được chiếm repo). Với văn bản nội bộ/hợp đồng, đây là rủi ro bảo mật; đề nghị **tắt tự cập nhật trong bản của bạn** hoặc ký số macro.
- Macro VBA cần bật macro/Trusted Location; nhiều môi trường doanh nghiệp chặn.
- Phát hành gần như hằng ngày → hành vi thay đổi nhanh, khó ổn định cho văn bản pháp lý.

## 4. Việc nên làm tiếp
1. Quyết định hướng A/B/C ở mục 1.
2. Nếu chọn A: dùng bộ tài liệu `01…06` + `CLAUDE.md` trong thư mục này.
3. Chuẩn bị 5–10 văn bản mẫu thật (đã ẩn thông tin) làm fixture — xem `docs/04-KIEM-THU.md`.

## 5. Kết quả kiểm chứng `.dotm` (2026-11-23, đã được người dùng cho phép)
Đã tải `ChuanHoa_VanBanDang_Win.dotm` (v2026.11.20.3, 180.156 byte; **sha256 khớp** `version.json`), chỉ giải nén zip và đọc `customUI14.xml`. **Không mở trong Word, không chạy macro, chưa giải mã VBA.**

Ribbon thực tế (tab "THỂ THỨC VĂN BẢN"):
- Nút **Chuẩn Hóa Thể Thức** (lề, phông, quốc hiệu/tiêu ngữ, số ký hiệu, ngày tháng, trích yếu, kính gửi, nơi nhận, chữ ký, rà soát logic), **Hoàn tác**, **Cập nhật** (tự cập nhật qua mạng).
- Chọn **Khối**: Đảng (HD 05-HD/VPTW), Chính quyền/Công an/Quân đội (NĐ 30/2020), Đoàn TNCS (HD 29), MTTQ và tổ chức CT-XH.
- **Tạo văn bản mẫu** theo khối (Nghị quyết, Quyết định, Kế hoạch, Báo cáo, Công văn, Tờ trình, Biên bản, Quy định, Quy chế…), **quản lý đơn vị ban hành**, **tùy chọn thể thức** (cỡ chữ, lề, thụt dòng, giãn dòng/đoạn), xuất/nhập cấu hình. Có nhắc "bản quyền" → có thể là sản phẩm có mã kích hoạt.

**Kết luận cập nhật:** add-in này nhắm **văn bản hành chính/Đảng–Đoàn–Nhà nước**, không có nút đánh số điều khoản đa cấp, mục lục, tham chiếu chéo hay hợp đồng. Mức tương thích với nhu cầu của bạn **thấp hơn ước tính ban đầu (~15–25%)**: chỉ trùng ở công văn và quy định/quy chế dạng mẫu. Nên giữ hướng A (viết mới); lấy từ repo ý tưởng: hoàn tác, tuỳ chọn trong khoảng cho phép, hồ sơ theo loại văn bản, xuất/nhập cấu hình.
Chưa biết: logic bên trong VBA (cần giải mã `vbaProject.bin`; chỉ làm nếu bạn muốn học thuật toán và đã chấp nhận rủi ro bản quyền).
