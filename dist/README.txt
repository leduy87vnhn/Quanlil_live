# Hướng dẫn sử dụng file exe với Google Sheet

1. Đặt file credentials.json vào cùng thư mục với file gui_fullscreen_match.exe (thư mục dist).
2. Khi nhập đường dẫn credentials trên giao diện, chỉ cần nhập: credentials.json
3. Nhập link Google Sheet và bấm tải dữ liệu.
4. Nếu vẫn không lấy được dữ liệu, gửi lại log để kiểm tra tiếp.

Lưu ý: Nếu muốn đóng gói credentials.json vào exe, cần sửa code để luôn lấy file từ thư mục tạm hoặc sys._MEIPASS.
