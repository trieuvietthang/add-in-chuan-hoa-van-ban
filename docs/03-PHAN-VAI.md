# 03 — Phân vai (Roles)

Một phiên Claude có thể **đóng lần lượt các vai** dưới đây (đội hình mặc định), hoặc giao từng vai cho subagent riêng khi người dùng yêu cầu. Mỗi vai có **đầu vào → đầu ra → cổng chuyển giao**. Không bỏ qua cổng.

```
Analyst → Architect → Tester(đỏ) → Developer → Tester(xanh) → Reviewer → Release
                                         ↑          │
                                         └── Fixer ←┘ (khi test/review đỏ)
```

## 1. Analyst — Phân tích nghiệp vụ
- **Việc:** biến yêu cầu thành kịch bản `KB-xx` (Cho/Khi/Thì), cập nhật `06-NGHIEP-VU`.
- **Vào:** yêu cầu người dùng, văn bản mẫu. **Ra:** kịch bản có tiêu chí chấp nhận đo được, danh sách câu hỏi mở.
- **Cổng:** kịch bản không còn từ mơ hồ ("đẹp", "chuẩn") — phải có số đo (cm, pt, chuỗi số).
- **Cấm:** viết code.

## 2. Architect — Thiết kế
- **Việc:** chọn module, hàm công khai, cấu trúc dữ liệu (ví dụ cấu hình bộ đánh số dạng bảng), ghi ADR nếu đổi quyết định.
- **Ra:** sơ đồ module ngắn + chữ ký hàm trong `src/…` (stub), cập nhật `07-ADR.md`.
- **Cổng:** mỗi kịch bản P0 ánh xạ được tới hàm cụ thể; logic thuần tách khỏi logic Word.

## 3. Tester — Kiểm thử
- **Việc:** viết fixture, golden, test (xem `04-KIEM-THU.md`); chạy; xác nhận **đỏ trước, xanh sau**.
- **Cổng vào Developer:** test đỏ vì đúng lý do (thiếu tính năng), không đỏ vì lỗi hạ tầng.
- **Cổng ra:** toàn bộ test xanh + test hồi quy cũ không vỡ.
- **Cấm:** sửa test cho khớp code sai; chỉ sửa test khi kịch bản đổi (và ghi lý do).

## 4. Developer — Lập trình
- **Việc:** viết code tối thiểu làm test xanh, theo `02-LUAT-CODE.md`; build bằng `tools/build.ps1`.
- **Ra:** commit nhỏ gắn `KB-xx`.
- **Cấm:** thêm tính năng ngoài kịch bản; refactor lẫn với sửa hành vi.

## 5. Fixer — Sửa lỗi
- **Việc:** chạy vòng lặp ở `05-VONG-LAP-SUA-LOI.md`: tái hiện → khoanh vùng → sửa gốc rễ → thêm test hồi quy.
- **Cổng:** mỗi lỗi có test hồi quy riêng; nêu rõ nguyên nhân gốc, không "vá triệu chứng".
- **Giới hạn:** 3 vòng cho 1 lỗi, quá thì dừng, báo người dùng kèm giả thuyết.

## 6. Reviewer — Rà soát
- **Việc:** duyệt theo checklist `02-LUAT-CODE.md` mục 9 + kiểm tra tiêu chí `01-KICH-BAN.md`; thử thủ công một văn bản thật nếu có Word.
- **Ra:** danh sách phát hiện (nặng/nhẹ). Nặng = chặn. Nên là một lượt đọc độc lập, không do chính người vừa viết code duyệt trong cùng một mạch nếu có thể dùng subagent.

## 7. Release — Đóng gói
- **Việc:** build sạch, chạy toàn bộ test, đổi số phiên bản `YYYY.MM.DD.N`, ghi changelog, ký số (nếu có chứng chỉ), tạo `dist/` + hướng dẫn cài.
- **Hỏi trước** khi: cài vào Word của người dùng, tạo tag/release, đẩy lên GitHub, xóa bản cũ.

## Quy ước bàn giao giữa các vai
Mỗi bàn giao ghi 4 dòng vào mô tả commit hoặc báo cáo: **Làm gì · Bằng chứng (lệnh + kết quả) · Rủi ro · Việc tiếp theo**.

## Phân vai người dùng ↔ Claude
- **Người dùng:** chốt nghiệp vụ (kiểu đánh số công ty), cung cấp văn bản mẫu ẩn danh, duyệt cuối, quyết định việc có tác động ngoài máy.
- **Claude:** tự phân tích–thiết kế–test–code–sửa lỗi trong vòng lặp; chỉ dừng hỏi ở những điểm trên.
