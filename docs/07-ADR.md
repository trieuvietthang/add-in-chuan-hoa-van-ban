# 07 — Nhật ký quyết định (ADR) & Changelog

Mỗi quyết định kiến trúc/nghiệp vụ lớn: 1 mục. Không xóa mục cũ, chỉ đánh dấu **Thay thế bởi ADR-xxx**.

## ADR-001 — Viết mới thay vì phát triển tiếp repo tham chiếu
- **Ngày:** 2026-11-23 · **Trạng thái:** Đề xuất (chờ người dùng xác nhận hướng A/B/C ở `00-DANH-GIA-TUONG-THICH.md`)
- **Bối cảnh:** `CuongBC195/add-in-office` chỉ có README + file phát hành; không có source/license.
- **Quyết định:** viết mới; repo chỉ làm tham chiếu chức năng.
- **Hệ quả:** tự chịu trách nhiệm toàn bộ; không rủi ro bản quyền nếu không sao chép.

## ADR-002 — Nền tảng: VBA `.dotm`, nguồn dạng text
- **Trạng thái:** Đề xuất
- **Lý do:** offline, tự test qua COM, tương đồng sản phẩm tham chiếu, đúng nhu cầu thao tác ListTemplate/Style.
- **Đánh đổi:** chính sách chặn macro; VBA kém công cụ; chỉ Windows tốt nhất. **Phương án khác:** Office.js web add-in (đa nền tảng, cần host HTTPS, API đánh số đa cấp hạn chế hơn). Đổi nền tảng cần ADR mới.

## ADR-003 — Đánh số bằng ListTemplate gắn Style; TOC bằng field
- **Trạng thái:** Chốt (luật cứng trong `CLAUDE.md`).

## ADR-004 — Chuỗi tiếng Việt qua `modStrings`
- **Trạng thái:** Đề xuất; chọn giữa `ChrW` và file tài nguyên UTF-8 khi có prototype đầu tiên.

## ADR-005 — Không tự cập nhật, không mạng trong add-in
- **Trạng thái:** Chốt. Lý do: văn bản pháp lý/nội bộ nhạy cảm; cơ chế cập nhật tự động của bản tham chiếu tin sha256 lấy từ cùng nguồn.

## ADR-006 — Kiến trúc theo "Hồ sơ" (Profile)
- **Ngày:** 2026-11-23 · **Trạng thái:** Đề xuất
- **Quyết định:** mọi quy ước (cấp, số, thụt, trang, mục lục, luật kiểm tra) nằm trong Profile (tệp văn bản, xuất/nhập được); mã chỉ đọc Profile. Lý do: 3 mẫu của bạn dùng 3 quy ước khác nhau (`08`); ý tưởng "Khối" + "Tùy chọn trong khoảng" + "Xuất/Nhập" lấy từ ribbon repo tham chiếu.

## ADR-007 — Tên style theo thói quen người dùng (`00Chuong…04Tiet`)
- **Trạng thái:** Chốt cho HopDong/QuyDinh (M1); Công văn dùng `CV1…CV6` (chốt M1). Thay thế tiền tố `VB-` ở `02-LUAT-CODE` mục 3/4 (các ví dụ `VB-` ở file đó/`06` mục 2–3 là cũ, xem mục 7 của `06`). Có bảng `legacyAliases` nhận diện tài liệu cũ; chuyển `02Khoản` → `02Khoan`.

## ADR-008 — Định dạng lưu Profile: INI UTF-8 (CHỐT 2026-11-23, M1)
- **Quyết định:** tệp `config/profiles/<Ten>.ini` cạnh tệp `.dotm` (build.ps1 sao chép), đọc lúc chạy bằng `ADODB.Stream` (UTF-8) và tự parse thành Dictionary (`modProfile`). Thông báo tiếng Việt ở `config/messages.vi.ini`, nên mọi chuỗi tiếng Việt nằm ngoài mã VBA (VBA chỉ ASCII). Khóa: `[profile] name,title,levels`; `[levelN] label,style,format,numstyle,numpos,textpos,trailing,bold,align,outline,resetonhigher,detect`.

## ADR-009 — Chính tả offline 4 tầng, không gọi mạng
- **Trạng thái:** Đề xuất. Không dựa Word proofing tiếng Việt (xác minh S5). Giới hạn đã nêu ở `09` mục 5.2.

## ADR-010 — Dẫn chiếu: REF field + bản đồ đổi số
- **Trạng thái:** Đề xuất; phụ thuộc S3. Phương án B nếu REF không cho chuỗi mong muốn: bookmark + macro cập nhật.

## ADR-011 — Dữ liệu thật của người dùng không vào repo
- **Trạng thái:** Chốt. Fixture tổng hợp; văn bản thật chỉ chạy trên bản sao tạm, báo cáo chỉ gồm số liệu cấu trúc.

## ADR-012 — Ribbon qua `customUI14.xml` chèn vào gói `.dotm` bằng `build.ps1`
- **Trạng thái:** Đề xuất; phụ thuộc S7.

## Kết quả spike M0
| Spike | Ngày | Kết luận |
|---|---|---|
| S1 | 2026-11-23 | **Đạt.** `ListTemplates.Add` + `LinkedStyle` dựng được 4 cấp gắn style 01Dieu…04Tiet; `ListString` ra `Điều 1.` / `1.1` / `a)` / `(i)`; Điều sau tự thành `Điều 2.`, `2.1`. Script: `tests/spikes/S1-listtemplate.ps1` (Word 16.0). Chưa so `numbering.xml` với F3. |
| S2 | 2026-11-23 | **Đạt (mức cơ bản).** Gắn lại `01Dieu` sang template Quy định: Điều ở cấp 3, đếm liên tục xuyên Chương (`ResetOnHigher=0`), `CHƯƠNG I/II` La Mã, Khoản `1.`. Cần xác nhận thêm khi chuyển hồ sơ trên tài liệu đã có nội dung (M1). |
| S3 | 2026-11-23 | **Đạt.** `REF bookmark \r`, `\n`, `\w` trên mức Điều (`Điều %1.`) cho `Điều 2` (**có tiền tố "Điều", bỏ dấu chấm cuối**) → khi chèn dẫn chiếu không viết thêm chữ "Điều". Mức Khoản (`%1.%2`) cho `2.1` ở cả `\r` và `\w`. `\h` trả nội dung đoạn (không dùng). Chưa thử Khoản kiểu `%n.` (Quy định) và dẫn chiếu Điểm `a)` kèm ngữ cảnh. |
| S4 | 2026-11-23 | **Đạt.** Đặt `Content.Font.Name/Size/Color` về giá trị chuẩn giữ nguyên in đậm (40/40 từ), 400 đoạn trong 0,21 s. `Font.Reset()` thì mất in đậm (0/40) → **không dùng Reset**. Lưu ý: cách này tạo định dạng trực tiếp trùng giá trị style; M2 cần tìm cách gỡ hẳn (thử đặt theo style/`ParagraphFormat.Reset` + khôi phục B/I/U theo run). |
| S5 | 2026-11-23 | **Kết luận: Word máy này KHÔNG có từ điển chính tả tiếng Việt** (`Languages(1066)` có tên nhưng `ActiveSpellingDictionary` rỗng). Chính tả phải tự làm (ADR-009 giữ nguyên). Tra tập âm tiết trong bộ nhớ: 10.000 từ ≈ 0,23 s (HashSet) — đủ nhanh; đọc UTF-8 đúng dấu. Chưa có danh sách âm tiết thật. |
| S6 | 2026-11-23 | **Đạt.** `UndoRecord.StartCustomRecord…End` bao đổi font toàn văn, đổi style (ListTemplate tự áp), đổi lề trang và chèn field: **một `Undo(1)`** khôi phục đúng trạng thái (so style|số|font|chữ từng đoạn, lề, field). Chưa thử với bảng/track changes/văn bản lớn. |
| S7 | 2026-11-23 | **Đạt (đã xác nhận hiển thị 2026-11-23: người dùng cài vào STARTUP, tab "Soạn thảo chuẩn" hiện, các nút chạy đúng mô tả).** Ban đầu đạt một phần: Tạo `.dotm` bằng COM (`SaveAs2(path, 15)`; 13 là .docm), nạp module VBA qua `VBProject` (AccessVBOM cần bật), chèn `customUI/customUI14.xml` + quan hệ vào `_rels/.rels` bằng zip, mở lại bình thường, macro còn nguyên. **Chưa kiểm chứng ribbon hiển thị** (cần nạp làm add-in ở STARTUP — ghi vào máy người dùng, phải hỏi trước). |
| S8 | 2026-11-23 | **Đạt.** Dấu phân cách danh sách trên máy này là `,` (`International(17)`; 5 không hợp lệ). `TOC \h \z \t "00Chuong,1,01Dieu,2"` cho đúng mục + số trang; `\t` đủ dùng, `\u` không gây hại. Lưu ý: đoạn chứa field TOC phải là style Normal (đoạn kế thừa style Điều sẽ bị tính/lỗi). |

**Bẫy kỹ thuật đã gặp (bắt buộc nhớ):**
1. Đường dẫn truyền vào COM phải là `[string]` thuần; `Join-Path` trả PSObject làm `SaveAs2` **treo vô hạn** (cả khi Word ẩn).
2. `AccessVBOM` bị Word ghi đè về 0 khi thoát nếu Word đang chạy lúc đặt giá trị → phải đặt khi Word đóng.
4. `MsgBox` của VBA là ANSI: chữ Việt có dấu ngoài bảng mã hệ thống hiện thành `?` (người dùng thấy ở M2). Thông báo trong `config/messages.vi.ini` viết **không dấu**; không đưa chữ trong văn bản người dùng vào MsgBox. Ribbon (XML) và nội dung văn bản không bị ảnh hưởng.
3. Test script luôn chạy trong tiến trình con có timeout; dọn tiến trình Word mồ côi theo PID đã xác định, không kill hàng loạt.

## ADR-013 — API mỏng chỉ dùng chuỗi/số; UndoRecord gắn tài liệu đích (M1)
- **Ngày:** 2026-11-23 · **Trạng thái:** Chốt
- **Bối cảnh:** `Application.Run` từ bên ngoài không truyền được đối tượng (Range/Document) → lỗi `Member not found`; và chỉ tìm module trong tài liệu đang *active* (add-in mở như tài liệu thường). `UndoRecord` gắn với tài liệu active tại thời điểm `Start`.
- **Quyết định:** lõi nhận `Range`/`Document` (`modApply`, `modConvert`); lớp `modApi` bọc bằng tham số chuỗi/số (`docName, firstPara, lastPara…`) cho test/tự động hóa. `modUndo.BeginUndo(doc, …)` kích hoạt tài liệu đích rồi mới bắt đầu ghi (ngoài thực tế tài liệu đang soạn đã active → không đổi gì). Test kích hoạt tài liệu add-in trước mỗi `Run`.
- **Ghi chú kỹ thuật khác:** lỗi `438` trong macro trả ra ngoài là `Member not found`; macro không tồn tại là `Can't run the specified macro`. `Step` là từ khóa VBA (không đặt tên biến). `ListTemplates.Add` phải gọi `Add(OutlineNumbered:=True)` rồi gán `.Name` (truyền `Name` vào `Add` gây lỗi 5). Với bullet phải gán `NumberFormat` trước `NumberStyle` rồi gán lại `NumberFormat`.

## ADR-014 — Chính tả tầng W5 bằng bảng cặp âm tiết từ ngữ liệu Apache-2.0 (M5, 2026-11-25)
- **Trạng thái:** Chốt. Bổ sung ADR-009.
- **Quyết định:** dùng bảng tần suất âm tiết/cặp âm tiết (`config/bigram.vi.txt`) sinh offline bằng `tools/build-bigrams.ps1` từ cột `correct_text` của `bmd1905/error-correction-vi` (Apache-2.0, ghi nguồn trong hướng dẫn). Cặp lạ + đổi dấu một âm tiết cho cặp phổ biến (`min_alt`) và âm tiết thay thế phổ biến hơn `min_ratio` lần → báo W5 kèm gợi ý.
- **Không dùng:** mô hình BARTpho (cần Python/mạng, trái ADR-005); VMDS (GPL-3.0, yêu cầu liên hệ tác giả); mô hình học máy bất kỳ trong add-in.
- **Đánh đổi:** không bắt câu vô nghĩa gồm từ đúng; ngữ liệu tin tức thiếu thuật ngữ pháp lý hiếm (xử lý bằng `[ignore]`); thêm 2,5 MB dữ liệu và ~1–2 s nạp lần đầu.

## ADR-015 — Cài đặt bằng script có hỏi, không ký số tự động (M6, 2026-11-25)
- **Trạng thái:** Chốt.
- **Quyết định:** `tools/install.ps1` (cài/gỡ, `-DryRun`, sao lưu config cũ, không chạm registry/cài đặt bảo mật), `tools/package.ps1` đóng gói zip. **Không tự ký số và không tự sửa Trusted Locations** (luật cứng 8: đổi cài đặt bảo mật phải do người dùng/IT). Hướng xử lý macro bị chặn ghi ở `docs/11`.
- **Chưa làm (đưa vào backlog):** ký số bằng chứng chỉ công ty.

## ADR-016 — Form "Sửa hồ sơ" là bộ soạn INI tổng quát dựng lúc chạy (M6, 2026-11-25)
- **Trạng thái:** Chốt.
- **Quyết định:** `frmProfile` dựng điều khiển bằng `Controls.Add` lúc chạy (mỗi khóa INI một dòng nhãn + ô nhập), nhãn/kiểu/chỉ-đọc/tên tệp lấy từ `config/ui.vi.ini`; ghi bằng `modIniEdit.SaveEdits` (giữ chú thích, sao lưu `.bak`, UTF-8 không BOM, kiểm tra giá trị trước khi ghi).
- **Lý do:** thêm khóa INI mới chỉ cần thêm dòng gợi ý, không phải vẽ lại form; Unicode hiển thị đúng trong form (MsgBox thì không).
- **Kỹ thuật:** `VBComponents.Import` file `.frm` lỗi `&H80004005` trong Word ẩn (kể cả form rỗng) → `build.ps1` tạo UserForm bằng `VBComponents.Add(3)`, đặt kích thước từ phần đầu `.frm` và nạp mã bằng `CodeModule.AddFromString`. Sự kiện Change của điều khiển dựng động có thể chưa chạy trong `UserForm_Initialize` → gọi nạp tệp đầu tiên tường minh.
- **Đánh đổi:** không sửa được cặp sai–đúng/nhóm biến thể (nhiều dòng, khóa tự do) bằng form; hồ sơ mới/số cấp đổi vẫn sửa tệp.

## Mẫu ADR
```
## ADR-00N — Tiêu đề
- Ngày · Trạng thái (Đề xuất/Chốt/Thay thế bởi …)
- Bối cảnh · Quyết định · Lý do · Đánh đổi · Hệ quả
```

## Changelog
| Phiên bản | Ngày | Thay đổi | Kịch bản |
|---|---|---|---|
| (chưa phát hành) | | Khởi tạo bộ tài liệu dự án | — |
| 0.1-dev (M1) | 2026-11-23 | Hồ sơ INI (HopDong/QuyDinh/CongVan), style + ListTemplate đa cấp idempotent, áp cấp, tăng/giảm cấp, chuyển số gõ tay (xem trước, 1 Ctrl+Z), ribbon (chưa kiểm hiển thị). 27 test xanh. | KB-A1–A4, A7 |
| 0.2-dev (M2) | 2026-11-23 | Chuẩn hóa trang/phông/giãn dòng/thụt đầu dòng/số trang theo chuẩn trong `config/standards` (ND30, ISO9001); xem trước + 1 Ctrl+Z + idempotent; nút ribbon "Chuẩn hóa trang & phông". 47 test xanh. | KB-B1–B3, F1 |
| 0.3-dev (M3) | 2026-11-23 | Mục lục theo hồ sơ (chèn/cập nhật, xử lý con trỏ trong bảng, hỏi chuyển số gõ tay); style khởi tạo từ Normal, chữa ngắt trang nhiễm; thông báo không dấu. | KB-C1 |
| 0.4-dev (M4) | 2026-11-23 | Kiểm tra dẫn chiếu (R1/R2/R4/R5), chuyển REF, cập nhật, chèn dẫn chiếu, nhảy tới lỗi. 79 test xanh. | KB-D2, A6 |
| 0.5-dev (M5) | 2026-11-25 | Chính tả offline S1–S6, W1–W5 (W5: bảng cặp âm tiết Apache-2.0). Vá: dẫn chiếu "Điều A đến Điều B của văn bản ngoài", W4 bỏ qua IN HOA, "qu"+ô. 98 test xanh, ribbon đã kiểm tay. | KB-D3 |
| 0.6.0 (M6) | 2026-11-25 | `install.ps1` (cài/gỡ/DryRun/sao lưu), `package.ps1` (zip), hướng dẫn `docs/11`. | KB-G |
| 0.6.1 (M6) | 2026-11-25 | Form "Sửa hồ sơ…" (nút trong nhóm Hồ sơ): sửa chuẩn NĐ30/ISO, hồ sơ đánh số, ngưỡng chính tả; kiểm tra giá trị, `.bak`. 114 test xanh. | KB-G |
| 0.6.2 | 2026-11-26 | Tăng tốc chính tả (KB-E1 chập chờn trên laptop: 3,7–10,8 s, ngưỡng 5 s): bảng bigram chuyển từ `Scripting.Dictionary` sang bảng băm mảng (nạp 2,0 → 0,4 s, tra 9 → 3 µs); ghi nhớ kết quả W3/W5 theo chữ trong một lần chạy; chuỗi đoạn văn truyền `ByRef`. 3.000 đoạn: lặp lại 5,4 → ~2 s, lần đầu 9–12 → ~6 s (máy bận). Thêm test `KB-E1 3000-doan-lan-dau`. Kết quả báo lỗi không đổi. 115 test xanh. | KB-E1, D3 |
