# Soạn thảo văn bản chuẩn — Add-in Word (Hợp đồng · Công văn · Quy định nội bộ)

Add-in VBA (`.dotm`) cho Microsoft Word trên Windows, chạy hoàn toàn offline. Giúp người làm văn bản Việt Nam:

- đánh số điều khoản đa cấp tự cập nhật (Chương/Mục/Điều/Khoản/Điểm/Tiết; công văn `I. 1. 1.1. a) (i)`);
- chuẩn hóa trang, phông, giãn dòng theo NĐ 30/2020 hoặc ISO 9001 (cấu hình được);
- chuyển số gõ tay thành đánh số tự động, tạo mục lục chuẩn;
- kiểm tra và chuyển dẫn chiếu "Điều N" thành dẫn chiếu tự động;
- kiểm tra chính tả offline (lỗi kiểu chữ, âm tiết không hợp lệ, cặp sai–đúng, nghi gõ sai dấu);
- sửa hồ sơ/chuẩn thể thức bằng form trong Word.

## Bắt đầu
- Hướng dẫn sử dụng: [`docs/11-HUONG-DAN-SU-DUNG.md`](docs/11-HUONG-DAN-SU-DUNG.md)
- Ký số macro: [`docs/12-KY-SO.md`](docs/12-KY-SO.md)
- Quy ước dự án, kiến trúc, kế hoạch: [`CLAUDE.md`](CLAUDE.md), thư mục [`docs/`](docs/)

## Build và kiểm thử (cần Word trên Windows, đã đóng Word đang dùng)
```powershell
powershell -ExecutionPolicy Bypass -File tools\build.ps1      # src/ -> dist\SoanThaoChuan.dotm
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1  # chạy test trên Word thật, văn bản tổng hợp
powershell -ExecutionPolicy Bypass -File tools\package.ps1    # đóng gói zip
```
Build cần bật `AccessVBOM` ở máy phát triển (xem `docs/04-KIEM-THU.md`). Mã nguồn VBA chỉ dùng ASCII; chuỗi tiếng Việt đặt trong `config/*.ini`.

## Lưu ý
- Repo **không chứa văn bản thật** của người dùng; test dùng văn bản tổng hợp.
- `.dotm` trong `dist/` là sản phẩm build, không đưa lên kho mã; phát hành qua Releases kèm chữ ký số.

## Nguồn dữ liệu
`config/bigram.vi.txt` là bảng tần suất âm tiết/cặp âm tiết thống kê từ phần câu đúng của bộ ngữ liệu [`bmd1905/error-correction-vi`](https://huggingface.co/datasets/bmd1905/error-correction-vi) (Hugging Face, giấy phép Apache-2.0), sinh bằng `tools/build-bigrams.ps1`. Bảng không chứa câu gốc.
