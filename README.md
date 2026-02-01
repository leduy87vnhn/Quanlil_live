Quanlil_live — Release README

Quick start

1. Place your Google service account `credentials.json` (with Sheets API access) beside the exe, e.g. in the same folder as `Quanlil_live.exe`.
2. Launch `Quanlil_live.exe`.

Notes

- UI: there's a persistent toolbar with `Tải bảng` (load), `Lưu bảng` (save state), and `Thoát` (exit).
- Shortcuts: `Esc` toggles fullscreen, `Ctrl+Q` saves state and exits.
- Google Sheets safety: the application writes per-cell using a safe batch update; by default the packaged app only updates columns AA..AL to avoid overwriting other sheet data.
- vMix: the app reads Input 1 via `http://<vmix-host>:8088/API/` and uses `Function=SetText` to push fields back to vMix.

Packaging

This release zip contains the one-file exe and the sample `vmix_input1_temp.csv`. Do NOT commit your real `credentials.json` into repositories — place it manually into the release folder before running.

