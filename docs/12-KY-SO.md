# Ký số add-in (macro VBA) để Word không chặn

**Vì sao cần:** nhiều máy công ty đặt chính sách "chỉ chạy macro đã ký bởi nhà phát hành tin cậy". Khi đó thẻ **Soạn thảo chuẩn** không hiện dù đã cài đúng. Ký số là cách chuẩn; không nên hạ mức bảo mật của Word.

**Ghi nhớ quan trọng**
- Chữ ký VBA nằm *trong* tệp `.dotm`. **Mỗi lần build lại là mất chữ ký** → luôn ký bước cuối cùng, sau `tools\build.ps1` và test xanh, rồi mới đóng gói.
- Add-in không tự ký, không tự tạo chứng chỉ, không tự thêm nhà phát hành tin cậy (CLAUDE.md luật 8). Các bước dưới đây do người có thẩm quyền làm.
- Chưa có công cụ dòng lệnh đáng tin để ký dự án VBA trên máy này (không có `signtool`/`msosip`); ký qua giao diện VBE là cách Microsoft hỗ trợ.

## A. Dành cho IT (chứng chỉ ký mã của công ty) — khuyến nghị
1. Lấy chứng chỉ **Code Signing** của công ty (hoặc CA nội bộ) cài vào kho cá nhân `CurrentUser\My` của máy ký.
2. Chạy `tools\build.ps1`, `tests\run-tests.ps1` (xanh). Sao chép `dist\SoanThaoChuan.dotm` ra thư mục ký, **mở trực tiếp bằng Word** (không dùng bản đang nằm trong STARTUP).
3. `Alt+F11` → *Tools* → *Digital Signature…* → *Choose…* → chọn chứng chỉ công ty → OK → lưu tệp (`Ctrl+S`), đóng Word.
4. Kiểm tra đã ký: `tools\package.ps1 -Dotm <đường dẫn tệp đã ký>` in "Chữ ký số VBA trong gói: CÓ" (kiểm tra có phần `vbaProjectSignature.bin` trong gói). `install.ps1` cũng in trạng thái này khi cài.
5. Phân phối gói zip đã đóng. IT đưa chứng chỉ công ty vào **Trusted Publishers** của máy người dùng bằng Group Policy (hoặc người dùng chọn "Tin cậy mọi tài liệu từ nhà phát hành này" lần đầu mở).

## B. Máy cá nhân, không có chứng chỉ công ty (chỉ dùng cho máy của bạn)
1. Chạy `SELFCERT.EXE` (có sẵn trong thư mục Office, ví dụ `C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE`) → đặt tên chứng chỉ → OK. Đây là chứng chỉ tự ký, **chỉ tin cậy trên máy này**.
2. Làm các bước A.2–A.3 với chứng chỉ vừa tạo.
3. Khi mở Word lần đầu sau khi cài, Word hỏi về nhà phát hành: chọn tin cậy. Máy khác sẽ vẫn chặn (cần cách A).
Lưu ý: bước này tạo định danh và đổi cài đặt tin cậy của Word trên máy bạn, nên phải do bạn tự quyết định làm.

## C. Danh mục kiểm tra trước khi phát hành
- [ ] Build xong, test xanh.
- [ ] Ký tệp bằng chứng chỉ (A hoặc B), mở lại tệp đã ký bằng Word thấy "Signed" trong VBE → Tools → Digital Signature.
- [ ] `package.ps1 -Dotm <tệp đã ký>` báo CÓ chữ ký.
- [ ] Thử cài gói trên một máy có chính sách macro nghiêm ngặt: thẻ **Soạn thảo chuẩn** hiện, không có thanh cảnh báo macro.
- [ ] Hết hạn chứng chỉ: ghi ngày hết hạn; ký lại và phát hành bản mới trước ngày đó.

## D. Gửi IT (mẫu)
> Cần ký số dự án VBA trong tệp `SoanThaoChuan.dotm` (add-in Word nội bộ, chạy offline, không gọi mạng/ghi registry) bằng chứng chỉ Code Signing của công ty, và đưa chứng chỉ vào Trusted Publishers cho máy người dùng. Mã nguồn có thể cung cấp để rà soát; quy ước bảo mật của add-in nêu ở `docs/02-LUAT-CODE.md` mục Bảo mật.
