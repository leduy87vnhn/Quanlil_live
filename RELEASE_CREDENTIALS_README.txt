KHÔNG ĐƯỢC ĐẨY (DO NOT COMMIT):

- File credentials Google (service account JSON) KHÔNG nên được commit vào repository.
- Đặt file credentials (ví dụ: credentials.json) trực tiếp vào thư mục giải nén trước khi chạy exe.

Hướng dẫn nhanh:
1) Tạo Service Account trong Google Cloud, bật Sheets API và tạo key JSON.
2) Tải file JSON về và đổi tên thành `credentials.json`.
3) Đặt `credentials.json` vào cùng thư mục với `Quanlil_live.exe`.
4) Chạy `Quanlil_live.exe` và chọn credentials nếu cần.

Nếu bạn muốn tôi đóng gói credentials vào exe, hãy xác nhận explicit (đảm bảo bạn hiểu rủi ro bảo mật).