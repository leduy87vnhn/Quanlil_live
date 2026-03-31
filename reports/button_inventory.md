
# Inventory nút và chức năng — Quanlil_live

Tập tin này liệt kê tất cả các nút (buttons / controls) chính trong ứng dụng, vị trí và callback tương ứng.

> File chính: `src/gui_fullscreen_match.py` (hầu hết các nút). Một số control demo ở `src/gui_match_selector.py`.

## Ghi chú
- Mô tả ngắn: `Nút text` | `File` | `Callback / Hàm xử lý` | `Mô tả`.
- Một số nút tạo từ dialog dùng cùng callback chung (ví dụ OK/Hủy trong listbox dialog).

## Danh sách nút (tổng hợp)

### src/gui_fullscreen_match.py

- `F11` / `Escape` | `src/gui_fullscreen_match.py` | `_toggle_fullscreen()` | Toggle pseudo-fullscreen cho cửa sổ chính.
- `Ctrl+q` | `src/gui_fullscreen_match.py` | `_on_close_and_save_state()` | Lưu trạng thái autosave và thoát.

- Dialog `Chọn chế độ hiển thị ảnh` (`ask_image_mode_listbox`):
  - `OK` | same file | `on_ok` (local) | Chọn chế độ ảnh cho ô preview.
  - `Hủy` | same file | `on_cancel` | Đóng dialog.

- Dialog `Chọn hiệu ứng chuyển logo` (`ask_logo_effect_listbox`): `OK` / `Hủy` (tương tự).

- Popup `Preview: Ghi tất cả lên Sheets` (`show_preview_write_popup`):
  - `Chọn tất cả` | same file | `select_all` | Chọn tất cả item trong treeview preview.
  - `Bỏ chọn` | same file | `deselect_all` | Bỏ chọn.
  - `Ghi tất cả` (xanh) | same file | `do_write_all` → `write_all_vmix_to_sheet_async()` (fallback sync) | Thực thi ghi batch lên Google Sheets.
  - `Đóng` | same file | dialog close | Đóng popup.

- Preview 3x3 (`open_preview_all`):
  - `Auto-refresh` (Checkbutton) | same file | n/a (variable `auto_var`) | Bật/tắt vòng làm mới.
  - `Đóng preview` | same file | `on_close()` | Đóng cửa sổ preview.
  - `Mặc định 1→1` | same file | `map_default_seq()` | Map hàng bảng theo thứ tự (giữ center rỗng).
  - `Tải Logo footer` | same file | `choose_footer_logo_ctrl()` | Chọn logo footer cho preview.
  - `Xóa Logo footer` | same file | `clear_footer_logo_ctrl()` | Xóa logo footer.
  - Trong preview, config ô có các nút: `Chọn từ bảng`, `Nhập URL vMix`, `Tải ảnh`, `Tải nhiều logo + hiệu ứng`, `Tải Logo footer`, `Xóa Logo footer`, `Xóa`, `Đóng` — callbacks: `set_vmix_from_row`, `set_vmix_from_url`, `set_image`, `set_logo_playlist`, `set_footer_logo`, `clear_footer_logo`, `clear_cell`.

- Dialog `Cấu hình Preview` (`open_preview_mapping_dialog`):
  - Per-cell controls (trong mỗi ô 1..9):
    - `Chọn Bàn` | same file | `make_show_ban_listbox(...)` | Chọn tên bàn từ bảng và gán mapping.
    - `Nhập URL vMix` | same file | `make_set_vmix_url(...)` | Gán URL cho ô.
    - `Tải ảnh` | same file | `make_set_image(...)` | Gán ảnh cho ô.
    - `Nhiều logo + FX` | same file | `make_set_logo_playlist(...)` | Gán playlist logo + effect.
    - `Xóa` | same file | `make_clear(...)` | Xóa mapping.
  - Logo footer controls: `Tải Logo`, `Xóa Logo` → `choose_footer_logo` / `clear_footer_logo`.
  - `Áp dụng và lưu` (xanh) | same file | `apply_and_close()` | Lưu `self._last_preview_meta`, set `self._preview_meta_user_configured = True`, autosave.
  - `Đóng` | same file | dialog close.

- Bottom bar / controls chính:
  - `Cập nhật` (số bàn) | same file | `populate_table()` | Rebuild table theo `ban_var`.
  - `Lọc` | same file | `populate_table()` | Áp dụng filter theo input.
  - `Preview tổng hợp` | same file | `open_preview_all()` | Mở preview 3x3.
  - `Cấu hình Preview` | same file | `open_preview_mapping_dialog()` | Mở dialog cấu hình.
  - `Lưu bảng` | same file | `save_table_to_csv()` | Lưu CSV của bảng.
  - `Tải bảng` | same file | `load_table_from_csv()` | Load CSV vào bảng.

- Per-row buttons (trong `populate_table`):
  - `Chọn Bàn` (`btn_ban`) | same file | `show_table_schedule_popup(...)` | Hiển thị/điều chỉnh lịch/bàn.
  - `Kết quả` (`btn_ketqua`) | same file | `on_btn_ketqua` → `_run_ketqua_logic_for_row(...)` | Chạy logic Kết quả cho hàng.
  - `Gửi` (`btn_gui`) | same file | `on_push_to_vmix` → `push_to_vmix(idx)` | Gửi dữ liệu dòng lên vMix; đổi màu trạng thái.
  - `Sửa` (`btn_sua`) | same file | `open_edit_popup(idx)` | Mở popup Sửa kết quả.

- Popup Sửa kết quả (`open_edit_popup`):
  - `Lấy từ vMix` | same file | `prefill_from_vmix()` | Prefill form từ vMix Input1.
  - `Gửi` | same file | `send_and_close()` → `send_to_vmix()` | Gửi các field lên vMix (SetText API).
  - `Hủy` | same file | popup destroy.

- Log viewer (`vmix_debug.log` window):
  - `Xoá log` | same file | `_clear_log()` | Xóa file log.
  - `Làm mới` | same file | `_refresh_log()` | Refresh nội dung hiển thị.

- Utility / state buttons:
  - `Save preview now` (hàm) | same file | `save_preview_now()` | Lưu snapshot preview vào autosave.
  - `Manual Load State` (hàm) | same file | `manual_load_state()` | Cho chọn file `ui_state.pkl` và apply.

### src/gui_match_selector.py (demo)

- `Gửi lên vMix` | `src/gui_match_selector.py` | `push_to_vmix()` | Hiển thị messagebox; (TODO: gửi thực tế).

## Hành động tiếp theo (tùy bạn chọn)
- Tôi đã lưu báo cáo này vào: [reports/button_inventory.md](reports/button_inventory.md)
- Nếu bạn muốn, tôi có thể:
  - xuất CSV hoặc JSON của danh sách này, hoặc
  - thêm cột `Line` (số dòng chính xác) và commit thay đổi, hoặc
  - tìm thêm controls trong `tools/` và `src/send_to_sheet` nếu cần.

---
Report generated automatically from code scan.
