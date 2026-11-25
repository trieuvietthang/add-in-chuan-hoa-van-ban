# 10 — Kế hoạch phát triển

Ước lượng theo "phiên làm việc" (một phiên ≈ 1 buổi làm việc tập trung của Claude + bạn duyệt). Mỗi mốc có **tiêu chí ra** (không đạt thì không sang mốc sau). Chưa có mã nào được viết.

## Tổng quan
```
M0 Nền tảng & Spike ─► M1 Hồ sơ + Đánh số ─► M2 Chuẩn hóa định dạng ─► M3 Mục lục
        │                                                  │
        └──────────────► M4 Kiểm tra dẫn chiếu ──► M5 Chính tả ──► M6 Hoàn thiện & Triển khai
```
| Mốc | Nội dung | Phiên | Kịch bản liên quan |
|---|---|---|---|
| M0 | Môi trường, khung repo, 8 spike | 2–3 | KB-E |
| M1 | Hồ sơ (3 loại) + style + list template + áp cấp/ tăng–giảm cấp + chuyển số gõ tay | 4–5 | A1–A4, A7 |
| M2 | Chuẩn hóa trang/phông/giãn dòng + nhận diện cấu trúc + xem trước + undo | 4–5 | B1–B4, KB-F1 |
| M3 | Mục lục theo hồ sơ | 2 | C1–C2 |
| M4 | Kiểm tra dẫn chiếu + chuyển REF + sửa theo bản đồ đổi số | 4 | D1, D2, A6 |
| M5 | Chính tả offline 4 tầng + từ điển | 3 | D3 |
| M6 | Sửa hồ sơ (UI), xuất/nhập, cài đặt, ký số, hướng dẫn, nghiệm thu trên 3 mẫu | 3 | E, G |

Thứ tự ưu tiên theo giá trị cho bạn: **M1 → M2 → M3** (giải quyết ngay việc tay hàng ngày) → **M4** (rủi ro pháp lý) → M5 → M6.

## M0 — Nền tảng & Spike (chặn mọi mốc sau)
Việc:
1. Khung thư mục theo `CLAUDE.md`; `tools/build.ps1`, `tools/inspect.ps1`, `tests/run-tests.ps1` (mới ở dạng khung).
2. Xin bạn đồng ý bật `AccessVBOM` ở máy dev (ghi cách tắt lại) — hoặc chọn phương án nhập tay.
3. **8 spike** (mỗi spike: câu hỏi → thí nghiệm nhỏ → kết luận ghi vào `07-ADR.md`):

| Spike | Câu hỏi | Thí nghiệm | Ảnh hưởng |
|---|---|---|---|
| S1 | Dựng ListTemplate 4–6 cấp gắn style, thụt/Tab đúng như F1/F3? | Tạo bằng code, xuất `numbering.xml`, so với F3 (cấu trúc) | M1 |
| S2 | Cùng một style có thể liên kết cấp khác nhau giữa profile (Điều = cấp 0 ở HĐ, cấp 1 ở QĐ)? | Chuyển profile trên cùng tài liệu | M1 |
| S3 | `REF \r/\w/\n` với `Điều %n.` trả chuỗi gì (dấu chấm cuối, tiền tố "Điều")? | Tạo bookmark + REF, đọc `.Result` | M4 |
| S4 | Xóa font/cỡ/màu trực tiếp nhưng giữ B/I/U/highlight: dùng API nào, tốc độ trên 400 đoạn? | Thử `Font.Reset` theo run, hoặc đặt lại thuộc tính từng thuộc tính | M2 |
| S5 | Word trên máy này có kiểm tra tiếng Việt? Cách nạp tập âm tiết (UTF-8) và tốc độ | `Languages(wdVietnamese)`; đọc tệp; đo 10.000 từ | M5 |
| S6 | `UndoRecord` bao ListTemplate/field/style? Backup tạm bằng `FormattedText` đủ tin cậy? | Chạy pipeline giả → `Undo` → so inspect | M2 |
| S7 | Chèn `customUI14.xml` vào `.dotm` sau khi lưu & vẫn mở được; tab hiển thị | Script zip + mở bằng COM | M6/M1 |
| S8 | TOC bằng `\t` danh sách style tùy biến + `\u`; dấu phân cách `;` hay `,` trên máy này | Tạo TOC trên tài liệu tổng hợp | M3 |

**Trạng thái M0 (2026-11-23): ĐÃ ĐẠT** — 8 spike có kết luận (xem `07-ADR.md`), `tools/build.ps1` + `tools/inspect.ps1` + `tests/run-tests.ps1` chạy xanh (4 PASS: Word khởi động, build .dotm có ribbon, macro `modStrings.Hello` chạy được, inspect ra JSON). Còn mở: ribbon hiển thị thật (cần cài STARTUP, hỏi trước).

Tiêu chí ra M0: 8 spike có kết luận; build.ps1 tạo được `.dotm` rỗng có ribbon, mở được trong Word; test chạy tự động 1 ca "xin chào".

## M1 — Hồ sơ & Đánh số
**Trạng thái (2026-11-23): XONG phần tự động** — KB-A1, A2, A3, A4, A7 + idempotent + Undo + đổi hồ sơ + lưu/mở lại: 27 test xanh. **Đã xác nhận:** ribbon hiển thị và các nút chạy đúng trong Word thật (người dùng kiểm tay, 2026-11-23). **Còn lại (cần bạn):** chốt câu hỏi mở 7–8 ở `06` (Khoản HĐ `1.1` hay `1.`; dãy ký hiệu công văn); xác nhận tên chương một dòng ("CHƯƠNG I tên") hay hai dòng.
- Hồ sơ `HopDong`, `QuyDinh`, `CongVan` (INI/JSON) + validate.
- `EnsureStyles`, `EnsureListTemplate` (idempotent).
- Nút áp cấp: Chương/Mục/Điều/Khoản/Điểm/Tiết cho vùng chọn; ↑/↓ cấp; phím tắt.
- Chuyển số gõ tay → list thật (có xem trước).
- Ra: KB-A1–A4, A7 xanh trên fixture tổng hợp; idempotent; Undo.

## M2 — Chuẩn hóa định dạng
**Trạng thái (2026-11-23): phần tự động xong, 47 test xanh.** Đã làm: `modNormalize` (xem trước + áp dụng, đếm thay đổi, 1 UndoRecord, idempotent), chuẩn `config/standards/ND30.ini` và `ISO9001.ini` (giá trị tạm), A4 + lề, font/cỡ cho Normal/Header/Footer/style cấp + từng đoạn sai (giữ in đậm, không dùng Font.Reset), giãn dòng, thụt đầu dòng Normal, số trang bằng field PAGE (không ghi đè nội dung header/footer), nút ribbon. **Chưa làm/chưa kiểm:** phát hiện section thừa (chỉ đếm khi cần), văn bản có track changes/nội dung bảo vệ, tài liệu lớn, kiểm tay ribbon M2; giá trị NĐ 30/ISO là tạm; bảng chịu giãn dòng/khoảng cách của style Normal.
- Nhận diện cấu trúc + bảng kế hoạch + xem trước.
- Trang (A4, lề), Normal (font/cỡ/giãn 1.15–1.2/thụt), header/footer số trang.
- Xóa định dạng trực tiếp thừa theo S4; phát hiện section thừa.
- Hoàn tác (Undo + dự phòng).
- Ra: nghiệm thu trên bản sao F1 và F3 (cấu trúc): 100% đoạn điều khoản đúng style; 0 run sai font/cỡ; không mất comment/footnote/bảng; chữ không đổi.

## M3 — Mục lục
**Trạng thái (2026-11-23): phần tự động xong, 58 test xanh.** `modToc`: TOC field theo `[toc] entries=` của hồ sơ, tiêu đề style riêng không vào mục lục, chèn lần 2 = cập nhật, 1 Ctrl+Z. Ribbon đã được người dùng kiểm tay trên văn bản thật (2026-11-23): chuẩn hóa ISO, chèn mục lục, đổi Điều rồi cập nhật đều đúng. Lỗi gặp và đã vá: `Styles.Add` chép định dạng đoạn đang chọn (ngắt trang trước) vào mọi style → style khởi tạo từ Normal, style cũ nhiễm được chữa khi áp cấp/chuẩn hóa; con trỏ trong bảng → chèn sau bảng; mục lục trống → hỏi chuyển số gõ tay. Phạm vi mục lục (Chương/Điều, bỏ Mục) là mặc định tạm. Chưa làm: tùy chọn "không số trang", chèn ở đầu văn bản (KB-C2).
- Chèn/cập nhật theo hồ sơ; style toc; loại tiêu đề mục lục khỏi mục lục.
- Ra: F1 có TOC liệt kê đủ Điều; F3 có TOC Chương–Điều đúng; Cập nhật sau khi chèn Điều mới.

## M4 — Dẫn chiếu
**Trạng thái (2026-11-23): phần tự động xong, 79 test xanh; ribbon M4 chưa kiểm tay.** `modRefs`: kiểm tra dẫn chiếu nội bộ "Điều N / khoản K Điều N / điểm x khoản K Điều N" (mẫu regex ở `config/refs.vi.ini`, mức dẫn chiếu khai báo bằng `refkind=` trong INI hồ sơ), luật R1 (đích không tồn tại), R2 (tên trong ngoặc trùng Điều khác), R4 (gõ tay), R5 (khoản/điểm không thuộc Điều); bỏ qua dẫn chiếu ra văn bản ngoài ("của Luật…"). Chuyển "Điều N" gõ tay hợp lệ thành REF field (bookmark ẩn `_STC_n` bắt đầu từ ký tự thứ 2 của đoạn để chèn Điều mới phía trên không kéo bookmark), cập nhật, chèn dẫn chiếu mới, nhảy tới lỗi; 1 Ctrl+Z. Số Điều đổi → REF tự theo. **Chưa làm:** R3 (tự dẫn chiếu), R6; dẫn chiếu "khoản K"/"Mục"/"Chương"/"Phụ lục" đứng riêng; chuyển REF cho khoản/điểm (chỉ phần "Điều N" thành REF); bản đồ đổi số cho dẫn chiếu gõ tay sai; dẫn chiếu dạng "Điều 5, 6 và 7"/"Điều 5 đến Điều 7" chỉ nhận từng "Điều N" riêng.
- Chỉ mục địa chỉ + trích dẫn chiếu + luật R1–R6 + báo cáo nhảy tới vị trí.
- Chuyển REF; sửa theo bản đồ đổi số.
- Ra: bộ test có dẫn chiếu cố ý sai/ đúng/ ngoài; tỷ lệ phát hiện 100% trên bộ lỗi chủ ý, ≤ 5% báo nhầm trên bản tổng hợp sạch.

## M5 — Chính tả
- Tầng 1–4; từ điển người dùng; báo cáo.
- Ra: bắt đủ lỗi trong bộ lỗi chủ ý; tốc độ ≤ 5 s / 200 trang.
- **Trạng thái (2026-11-25):** phần tự động xong, 97 test xanh. Mã lỗi S1–S6/W1–W4 (`docs/01` KB-D3); âm tiết kiểm bằng luật ngữ âm (không cần danh sách từ); 3.000 đoạn ≈ 1 s; trên bản sao quy chế thật chỉ báo 7 mục (hai dấu cách/lặp từ), không báo nhầm âm tiết. Chưa làm: chú thích cuối trang/đầu-cuối trang, nhất quán số tiền–bằng chữ, từ điển người dùng qua giao diện. Ribbon chưa kiểm tay.
- **Cập nhật M5 (2026-11-25, sau kiểm tay):** vá W4 (bỏ qua chữ IN HOA như "BÊN B") và luật "qu"+ô ("quốc"). Thêm **W5 – cặp âm tiết lạ** (nghi gõ sai dấu, vd "quy đinh", "hợp đòng"): bảng `config/bigram.vi.txt` (148 nghìn dòng, 2,4 MB) sinh bằng `tools/build-bigrams.ps1` từ cột `correct_text` của bộ ngữ liệu `bmd1905/error-correction-vi` (Apache-2.0, cần ghi nguồn). Quy tắc: cặp không có trong bảng + đổi dấu một âm tiết cho cặp phổ biến (≥ `min_alt`) và âm tiết thay thế phổ biến hơn ≥ `min_ratio` lần → báo kèm gợi ý. Ngưỡng ở `[w5]` trong `config/spell.vi.ini`. 98 test xanh; bản sao quy chế thật: 0 báo W5. Giới hạn: ngữ liệu là tin tức/hành chính nên thiếu thuật ngữ pháp lý hiếm (báo nhầm thì thêm vào `[ignore]`); không bắt câu vô nghĩa gồm từ đúng; không bắt lỗi dấu khi cặp thay thế hiếm. Không dùng VMDS (GPL-3.0, cần liên hệ tác giả).

## M6 — Hoàn thiện
- UI sửa hồ sơ; xuất/nhập; ký số; cài đặt (`install.ps1` hỏi trước khi ghi STARTUP); hướng dẫn tiếng Việt; changelog; nghiệm thu trên 3 mẫu thật (bạn chạy trên bản sao).
- Ra: danh mục T4 thủ công đạt; bạn ký duyệt.
- **Trạng thái (2026-11-25):** xong `tools/install.ps1` (đã thử trên thư mục tạm: DryRun/cài/cài lại có sao lưu/gỡ), `tools/package.ps1` (zip 1,0 MB, thử giải nén + DryRun), `docs/11-HUONG-DAN-SU-DUNG.md`, ADR-014/015, changelog 0.6.0. Form "Sửa hồ sơ…" (`frmProfile`, `modIniEdit`, `config/ui.vi.ini`, ADR-016): 114 test xanh, đã cài và kiểm tay giao diện thật 2026-11-25 ("ok tốt"). Ký số: đã có `docs/12-KY-SO.md` (quy trình IT / SelfCert, danh mục phát hành) và `install.ps1`/`package.ps1` báo tệp có/chưa có chữ ký VBA; đã tự ký bằng SelfCert trên máy tác giả và xác nhận thẻ hiện, không cảnh báo macro (2026-11-25); máy khác vẫn cần chứng chỉ công ty. **Chưa làm:** ký bằng chứng chỉ công ty (cần chứng chỉ công ty/IT), nghiệm thu T4 trên 3 văn bản thật do bạn chạy trên bản sao, gói chưa thử trên máy sạch khác.

## Phân công (xem `03-PHAN-VAI.md`)
| Mốc | Analyst | Architect | Tester | Developer | Reviewer |
|---|---|---|---|---|---|
| M0 | xác nhận câu hỏi mở | chốt ADR | khung test | spike | — |
| M1–M5 | kịch bản chi tiết | thiết kế module | fixture + golden | code | rà soát độc lập |
| M6 | hướng dẫn | đóng gói | T4 | script cài | duyệt cuối |

## Quản lý rủi ro kế hoạch
- Nếu spike S1/S2 thất bại (không gắn style–list như mong muốn) → quay lại ADR-002, cân nhắc COM add-in/Open XML SDK (cần cài .NET) **trước** khi viết M1.
- Nếu `AccessVBOM` không được phép → dev loop chậm; lùi M1 +2 phiên hoặc dùng phương án thủ công.
- Phạm vi cuộn (scope creep): tính năng ngoài danh sách (tạo văn bản mẫu, tự cập nhật, Excel, Mac) → chỉ qua ADR mới.

## Việc cần bạn quyết/ cung cấp trước M0
1. Đồng ý bật `AccessVBOM` ở máy dev (tắt lại được)? 
2. Trả lời câu hỏi mở `06-NGHIEP-VU` mục 6 (đặc biệt: kiểu số Khoản hợp đồng, dãy ký hiệu công văn, lề ISO).
3. Đồng ý cách xử lý văn bản thật (chỉ trên bản sao tạm; không vào repo).
4. Phiên bản Word & có chính sách chặn macro/ yêu cầu ký số không?
