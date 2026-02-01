Bản phát hành: Quanlil_live (Windows exe, one-file)

Bao gồm trong ZIP:
- Quanlil_live.exe (từ thư mục dist)
- README.txt (hướng dẫn nhanh từ dist)
- ui_state.pkl (nếu có) - chứa trạng thái UI đã lưu
- RELEASE_CREDENTIALS_README.txt - hướng dẫn đặt credentials an toàn

Lưu ý:
- Không bao gồm file credentials JSON trong gói này vì lý do bảo mật.
- Trước khi chạy exe, hãy đặt file `credentials.json` (Google Service Account) vào cùng thư mục với exe.

Cách chạy (Windows):
- Giải nén `release_Quanlil_live.zip` vào một thư mục.
- Đặt `credentials.json` vào thư mục nếu bạn dùng tính năng Google Sheets.
- Chạy `Quanlil_live.exe` (double-click).

Muốn tôi đóng gói thêm file credentials (bạn chịu trách nhiệm về bảo mật)? Trả lời có/không.