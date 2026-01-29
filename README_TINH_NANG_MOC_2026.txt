# TỔNG HỢP TÍNH NĂNG & THÔNG SỐ CHI TIẾT PHẦN MỀM QUẢN LÝ LIVESTREAM (CỘT MỐC 18/01/2026)

## 1. Chức năng chính
- Quản lý bảng điểm các trận đấu bi-a, hiển thị và cập nhật trực tiếp trên giao diện.
- Kết nối, lấy dữ liệu thật từ Google Sheet (sheet 'Kết quả') qua tài khoản dịch vụ.
- Cho phép chọn file credentials Google API động từ giao diện.
- Tự động nhận diện và ánh xạ các cột: Trận, Số bàn, VĐV A, VĐV B.
- Cho phép nhập, chỉnh sửa, lưu và tải lại bảng điểm từ file CSV.
- Tự động cập nhật tên VĐV khi nhập số trận nếu có dữ liệu Google Sheet.
- Gửi dữ liệu từng dòng lên vMix qua HTTP API (tên, điểm, thời gian, địa điểm, chữ chạy, ...).
- Có thể mở link vMix từng bàn để chỉnh sửa trực tiếp.
- Có chức năng lọc bàn, tìm kiếm VĐV nhanh.
- Có thể thay đổi số lượng bàn hiển thị động.
- Có nút "Preview tổng hợp" để xem nhanh toàn bộ bảng điểm các bàn theo dạng lưới.
- Preview tổng hợp lấy dữ liệu trực tiếp từ vMix Input 1 từng bàn, hiển thị đầy đủ điểm số, AVG, lượt cơ, HR1, HR2 cho từng VĐV.
- Có nút "Lấy dữ liệu từ vMix" và checkbox "Tự động cập nhật (5s)" trong preview tổng hợp.
- Giao diện tối ưu cho màn hình lớn, màu sắc rõ ràng, dễ nhìn.

## 2. Thông số chi tiết giao diện
- Cột "Trận": width=16, font lớn, căn giữa, nhỏ gọn.
- Cột "Số bàn": width=28, font lớn, rộng gấp đôi cũ.
- Cột "Tên VĐV A/B": font đậm, readonly, luôn màu chữ đen.
- Cột "Điểm số": font lớn, căn giữa.
- Cột "Địa chỉ vMix": mặc định http://127.0.0.1:8088, có thể sửa từng dòng.
- Nút "Gửi" màu đỏ, nổi bật, gửi dữ liệu lên vMix.
- Nút "Sửa" mở link vMix.
- Cột trạng thái: hiển thị trạng thái gửi dữ liệu, màu xanh/đỏ.
- Header có các trường: Tên giải, Thời gian, Điểm số (đa dòng), Địa điểm, Chữ chạy.
- Có thể lưu/khôi phục toàn bộ bảng điểm và thông tin header ra file CSV.
- Có nút "Tải/Làm mới dữ liệu" (đã gộp).

## 3. Preview tổng hợp bảng điểm
- Hiển thị dạng lưới, mỗi bàn 1 khung.
- Lấy dữ liệu trực tiếp từ vMix Input 1 từng bàn.
- Hiển thị: Tên bàn, Trận, Tên 2 VĐV, điểm số lớn, AVG, lượt cơ, HR1, HR2 từng VĐV (ưu tiên lấy riêng từng VĐV, nếu không có thì lấy chung).
- Có nút "Lấy dữ liệu từ vMix" và checkbox "Tự động cập nhật (5s)".
- Khi đóng preview sẽ tự động dừng cập nhật.

## 4. Kết nối & xử lý dữ liệu
- Google Sheet: dùng GSheetClient, credentials động, lấy tối đa 2000 dòng.
- vMix: gửi/nhận dữ liệu qua HTTP API, lấy XML Input 1 từng bàn.
- Tự động nhận diện cột dữ liệu linh hoạt (không phụ thuộc tên cột tuyệt đối).
- Khi nhập số trận, tự động điền tên VĐV nếu có dữ liệu.
- Khi gửi lên vMix, gửi đầy đủ các trường cần thiết cho Input 1, backdrop, ketqua, chay chu.

## 5. Tính năng phụ trợ
- Lọc bàn, tìm kiếm VĐV nhanh.
- Có thể thay đổi số lượng bàn động.
- Lưu/khôi phục toàn bộ trạng thái bảng điểm, header ra file CSV.
- Có thể mở nhanh link Google Sheet, link vMix từ giao diện.
- Giao diện tối ưu cho livestream, font lớn, màu sắc rõ ràng.

## 6. Cột mốc lưu trữ
- File này là bản mô tả tính năng, thông số chi tiết tại thời điểm 18/01/2026.
- Khi phát triển tính năng mới hoặc sửa lỗi, luôn so sánh lại với file này để đảm bảo không làm mất các chức năng gốc.
- Nếu có lỗi phát sinh sau này, có thể dùng file này để đối chiếu, phục hồi hoặc rollback tính năng.

---
Người lưu: AI (GitHub Copilot)
Ngày: 18/01/2026
