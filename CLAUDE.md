# CLAUDE.md — Add-in Word "Soạn thảo văn bản chuẩn" (Hợp đồng · Công văn · Quy định nội bộ)

Đọc file này đầu mỗi phiên. Đây là "hiến pháp" của dự án; các file chi tiết nằm trong `docs/`.

## 1. Mục tiêu
Add-in cho **Microsoft Word (Windows trước)** giúp người dùng văn phòng Việt Nam:
1. Đánh số điều khoản nhanh, đúng chuẩn, **tự cập nhật** (Chương/Mục/Điều/Khoản/Điểm).
2. Định dạng chuẩn: font, lề, giãn dòng, thụt dòng, căn lề theo thể loại văn bản.
3. Sửa số thứ tự sai/gõ tay, đồng bộ tham chiếu chéo.
4. Xuất **mục lục** (TOC field) chuẩn.
5. Chuẩn hóa văn bản dán từ nguồn khác (xóa định dạng rác, khoảng trắng thừa).

Ngoài phạm vi (đến khi có lệnh mới): Excel, Mac, AI/LLM trong add-in, soạn nội dung pháp lý, tạo văn bản mẫu, tự cập nhật.

**Người dùng & nhu cầu thực tế (2026-11-23):** chuyên viên pháp chế soạn Hợp đồng, công văn, quy định nội bộ; hiện làm tay: (i) kiểm tra chính tả + dẫn chiếu sai; (ii) đánh số (công văn: `I. 1. 1.1. a) (i) -`; hợp đồng: Điều/Khoản/Điểm/Tiết); (iii) Phần/Chương/Mục/Điều/Khoản/Điểm/Tiết cho quy định; (iv) mục lục; (v) đồng nhất font/lề theo NĐ 30 hoặc ISO 9001 (TNR 12–13, multiple 1.15–1.2). Phương pháp hiện tại: tạo style 00Chuong…04Tiet + Multilevel List. Phân tích chi tiết: `docs/08`; giải pháp: `docs/09`; kế hoạch: `docs/10`. **Trạng thái: M0 và M1 (phần tự động) xong: hồ sơ INI, đánh số đa cấp, tăng/giảm cấp, chuyển số gõ tay, ribbon (27 test xanh); ribbon đã kiểm tay trong Word thật. M2 (chuẩn hóa trang/phông/giãn dòng/số trang theo `config/standards/*.ini`) xong phần tự động, 47 test xanh, ribbon M2 đã kiểm tay (thông báo chuyển sang tiếng Việt không dấu vì MsgBox ANSI). M3 (mục lục theo hồ sơ) xong phần tự động, 58 test xanh, ribbon M3 đã kiểm tay trên file thật (đã vá: ngắt trang nhiễm vào style, con trỏ trong bảng, mục lục trống); M4 (kiểm tra/chuyển REF dẫn chiếu "Điều N") xong phần tự động, 80 test xanh, ribbon M4 đã kiểm tay trên file thật (đã vá: lỗi biên dịch do khai báo cấp module sau thủ tục -> build có lint; chuỗi "Điều A đến Điều B của văn bản ngoài"); M5 (chính tả offline S1–S6, W1–W5 theo `config/spell.vi.ini` + `config/bigram.vi.txt` từ ngữ liệu Apache-2.0 bmd1905/error-correction-vi) xong phần tự động, 98 test xanh; đã vá 2 báo nhầm sau kiểm tay đầu; ribbon M5 gồm W5 đã cài và kiểm tay trong Word thật (2026-11-25, "đã chuẩn"); M6 (2026-11-25): xong `tools/install.ps1`, `tools/package.ps1`, `docs/11`; form "Sửa hồ sơ…" xong, đã cài và kiểm tay (114 test xanh); ký số tự ký trên máy tác giả xong (2026-11-25; máy khác cần chứng chỉ công ty, `docs/12`); còn nghiệm thu 3 mẫu thật.**

## 2. Bối cảnh
Repo tham chiếu `CuongBC195/add-in-office` **chỉ có bản phát hành, không có source**, nên dự án này **viết mới**. Xem `docs/00-DANH-GIA-TUONG-THICH.md`. Không copy code/binary của repo đó vào dự án khi chưa có quyền.

## 3. Quyết định kiến trúc (mặc định, đổi qua ADR trong `docs/07-ADR.md` nếu cần)
- **VBA, template `.dotm`**, nguồn là file text trong `src/` (`.bas`, `.cls`, `.frm`). `.dotm` trong `dist/` là **sản phẩm build**, không sửa tay.
- Lý do: chạy offline, giống sản phẩm tham chiếu, tự test được qua COM từ PowerShell. Đánh đổi: bị chính sách chặn macro → xem `docs/02-LUAT-CODE.md` mục Bảo mật; hướng nâng cấp sau: Office.js (cần ADR).
- Đánh số **bắt buộc** dùng Word **ListTemplate đa cấp gắn với Style** (không gõ số bằng tay, không dùng `Selection` để chèn số).
- Mục lục **bắt buộc** là TOC field dựa trên danh sách style của Hồ sơ (`\t`) + outline level (`\u`).
- Kiến trúc theo **Hồ sơ (Profile)**: Hợp đồng / Quy định nội bộ / Công văn × chuẩn trang NĐ30 / ISO9001 (ADR-006). Tên style theo thói quen người dùng `00Chuong…04Tiet` (ADR-007, thay tiền tố `VB-` ở các file cũ khi chốt).
- Ribbon cần `customUI14.xml` chèn vào gói `.dotm` bởi `tools/build.ps1` (ADR-012).
- Mỗi mốc chỉ bắt đầu khi 8 spike M0 đã có kết luận (`docs/10`).

## 4. Cấu trúc thư mục
```
CLAUDE.md
docs/            tài liệu (00…07)
src/             mã VBA dạng text (modXxx.bas, clsXxx.cls)
  core/          style, list template, field, tiện ích (không phụ thuộc UI)
  features/      từng tính năng (đánh số, định dạng, mục lục, chuẩn hóa)
  ui/            ribbon XML + callback
tests/
  fixtures/      .docx đầu vào (đã ẩn danh)
  golden/        kết quả mong đợi (.json mô tả cấu trúc, KHÔNG so byte)
  run-tests.ps1  chạy tất cả
tools/
  build.ps1      src/ -> dist/*.dotm (qua COM Word)
  export.ps1     dotm -> src/ (khi sửa trong VBE)
  inspect.ps1    dump cấu trúc docx ra JSON để so sánh
dist/            .dotm build ra (gitignore)
```

## 5. Quy trình làm việc (tóm tắt — chi tiết `docs/03-PHAN-VAI.md`, `docs/05-VONG-LAP-SUA-LOI.md`)
1. Nhận yêu cầu → đối chiếu `docs/01-KICH-BAN.md`; thiếu kịch bản thì **viết kịch bản + tiêu chí chấp nhận trước**.
2. Viết/cập nhật test (fixture + golden) → thấy **đỏ**.
3. Viết code theo `docs/02-LUAT-CODE.md` → `tools/build.ps1` → `tests/run-tests.ps1` → **xanh**.
4. Lỗi → vòng lặp sửa lỗi (tối đa 3 vòng, rồi dừng và báo cáo).
5. Hoàn tất: tự review theo checklist, cập nhật `CHANGELOG` trong `docs/07-ADR.md` nếu đổi quyết định.

## 6. Luật cứng (không được vi phạm)
1. **Không bao giờ sửa file gốc của người dùng.** Test chỉ chạy trên bản sao trong thư mục tạm.
2. Mọi thao tác ghi vào văn bản phải nằm trong **một `UndoRecord`** để Ctrl+Z hoàn tác một lần.
3. Không dùng `Selection` trừ lúc đọc vị trí con trỏ ban đầu; dùng `Range`/`Paragraph`.
4. Không gõ số thứ tự thành text ("Điều 1.", "1.1") — dùng ListTemplate. Có chuyển đổi số gõ tay → list thật là tính năng riêng.
5. Không gọi mạng, không tự cập nhật, không ghi registry, không chạy tiến trình ngoài trong add-in (xem Bảo mật).
6. Không đổi nội dung chữ của người dùng (trừ tính năng "chuẩn hóa" có xem trước/xác nhận).
7. Tiếng Việt Unicode (UTF-8 ở file md/json; VBA lưu ANSI → mọi chuỗi tiếng Việt trong VBA đặt trong `modStrings` dùng `ChrW`, xem `docs/02-LUAT-CODE.md`).
8. Hành động khó hoàn tác/ngoài máy (xóa file, push, gửi, cài vào Word thật của người dùng, **bật AccessVBOM/đổi cài đặt bảo mật Word**) → **hỏi trước**.
9. **Văn bản thật của người dùng là tài liệu mật**: chỉ đọc cấu trúc (style/numbering/trang/field), không chép nội dung vào md/log/fixture/báo cáo; chạy thử chỉ trên bản sao ở thư mục tạm. Fixture trong repo là bản tổng hợp (ADR-011).
10. Quy ước (Hồ sơ) đọc từ tệp cấu hình, **không hard-code** số liệu thể thức trong mã (ADR-006).

## 7. Định nghĩa "xong" (Definition of Done)
- [ ] Có kịch bản trong `01-KICH-BAN.md` và test tương ứng.
- [ ] `tests/run-tests.ps1` xanh, không test nào bị bỏ qua mà không ghi lý do.
- [ ] Hoàn tác 1 lần (Ctrl+Z) đưa văn bản về đúng trạng thái trước.
- [ ] Chạy 2 lần liên tiếp cho kết quả như nhau (idempotent).
- [ ] Không lỗi trên văn bản có bảng, header/footer, track changes, nội dung được bảo vệ.
- [ ] Đã tự review theo `docs/02-LUAT-CODE.md` mục 9.

## 8. Cách giao tiếp với người dùng
- Trả lời bằng **tiếng Việt**, ngắn, nêu rõ: đã làm gì / kiểm chứng ra sao / còn thiếu gì.
- Không khẳng định "chạy tốt" khi chưa chạy test; nói rõ phần nào chưa kiểm chứng (ví dụ chưa có Word để chạy).
- Gặp quyết định nghiệp vụ (kiểu đánh số hợp đồng của công ty bạn, lề chuẩn nội bộ) → hỏi 1 lần, ghi vào `docs/06-NGHIEP-VU-THE-THUC.md`.

## 9. Mục lục tài liệu
| File | Nội dung |
|---|---|
| `docs/00-DANH-GIA-TUONG-THICH.md` | Đánh giá repo tham chiếu so với nhu cầu |
| `docs/01-KICH-BAN.md` | Kịch bản người dùng + tiêu chí chấp nhận (nguồn sự thật của yêu cầu) |
| `docs/02-LUAT-CODE.md` | Luật viết code, cấu trúc, bảo mật, lỗi |
| `docs/03-PHAN-VAI.md` | Phân vai (Analyst, Architect, Dev, Tester, Fixer, Reviewer, Release) |
| `docs/04-KIEM-THU.md` | Chiến lược test tự động, fixture, golden, script |
| `docs/05-VONG-LAP-SUA-LOI.md` | Quy trình tự phát hiện–sửa lỗi |
| `docs/06-NGHIEP-VU-THE-THUC.md` | Quy tắc đánh số/thể thức Việt Nam, bảng thông số |
| `docs/07-ADR.md` | Nhật ký quyết định + changelog |
| `docs/08-PHAN-TICH-MAU-VAN-BAN.md` | Phân tích cấu trúc 3 mẫu thật: cách làm hiện tại, vấn đề P1–P12 |
| `docs/09-GIAI-PHAP-KIEN-TRUC.md` | Thiết kế: Hồ sơ, module, pipeline chuẩn hóa, đánh số, kiểm tra, mục lục |
| `docs/10-KE-HOACH-PHAT-TRIEN.md` | Mốc M0–M6, 8 spike, tiêu chí ra, rủi ro, việc cần bạn quyết |
| `docs/11-HUONG-DAN-SU-DUNG.md` | Hướng dẫn người dùng: cài đặt, lệnh, chính tả, tùy biến config, sự cố |
| `docs/12-KY-SO.md` | Ký số macro: quy trình cho IT / tự ký trên máy cá nhân, danh mục trước khi phát hành |
