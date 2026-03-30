import sys, os
import atexit
import tkinter as tk
from tkinter import ttk, messagebox


_DPI_AWARENESS_SET = False


def _enable_high_dpi_awareness():
    """Enable best-effort per-monitor DPI awareness on Windows before creating Tk."""
    global _DPI_AWARENESS_SET
    if _DPI_AWARENESS_SET or os.name != 'nt':
        return

    try:
        import ctypes
        user32 = ctypes.windll.user32

        # Windows 10+: prefer Per-Monitor V2
        try:
            if hasattr(user32, 'SetProcessDpiAwarenessContext'):
                dpi_context_per_monitor_v2 = ctypes.c_void_p(-4)
                if user32.SetProcessDpiAwarenessContext(dpi_context_per_monitor_v2):
                    _DPI_AWARENESS_SET = True
                    return
        except Exception:
            pass

        # Windows 8.1 fallback
        try:
            shcore = ctypes.windll.shcore
            process_per_monitor_dpi_aware = 2
            if shcore.SetProcessDpiAwareness(process_per_monitor_dpi_aware) == 0:
                _DPI_AWARENESS_SET = True
                return
        except Exception:
            pass

        # Legacy fallback
        try:
            if hasattr(user32, 'SetProcessDPIAware'):
                user32.SetProcessDPIAware()
                _DPI_AWARENESS_SET = True
                return
        except Exception:
            pass
    except Exception:
        pass

    _DPI_AWARENESS_SET = True


# Helper to fetch rows from Google Sheets using GSheetClient
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from src.gsheet_client import GSheetClient as _GSheetClient

# Public alias used by the module
GSheetClient = _GSheetClient

def ask_image_mode_listbox(parent=None, initial='fit'):
    modes = [
        ('fit', 'Fit (giữ tỉ lệ)'),
        ('cover', 'Fill/Cover (lấp đầy, có cắt)'),
        ('center', 'Center (căn giữa)')
    ]
    default_mode = initial if initial in ('center', 'fit', 'cover') else 'fit'
    selected = {'value': None}

    dlg = tk.Toplevel(parent) if parent is not None else tk.Toplevel()
    dlg.title('Chọn chế độ hiển thị ảnh')
    try:
        if parent is not None:
            dlg.transient(parent)
        dlg.grab_set()
    except Exception:
        pass

    tk.Label(dlg, text='Chọn chế độ hiển thị:', font=('Segoe UI', 11, 'bold')).pack(padx=12, pady=(10, 6))
    lb = tk.Listbox(dlg, height=len(modes), width=36, exportselection=False, font=('Segoe UI', 10))
    for key, label in modes:
        lb.insert('end', f'{key} - {label}')
    lb.pack(padx=12, pady=6)

    idx_default = 0
    for idx, (key, _) in enumerate(modes):
        if key == default_mode:
            idx_default = idx
            break
    try:
        lb.selection_clear(0, 'end')
        lb.selection_set(idx_default)
        lb.activate(idx_default)
        lb.see(idx_default)
    except Exception:
        pass

    def on_ok(event=None):
        try:
            cur = lb.curselection()
            if cur:
                selected['value'] = modes[cur[0]][0]
            else:
                selected['value'] = default_mode
        except Exception:
            selected['value'] = default_mode
        try:
            dlg.destroy()
        except Exception:
            pass

    def on_cancel(event=None):
        selected['value'] = None
        try:
            dlg.destroy()
        except Exception:
            pass

    btns = tk.Frame(dlg)
    btns.pack(padx=12, pady=(6, 10))
    tk.Button(btns, text='OK', width=12, command=on_ok).pack(side='left', padx=6)
    tk.Button(btns, text='Hủy', width=12, command=on_cancel).pack(side='left', padx=6)

    lb.bind('<Double-Button-1>', on_ok)
    dlg.bind('<Return>', on_ok)
    dlg.bind('<Escape>', on_cancel)

    try:
        dlg.wait_window()
    except Exception:
        pass
    return selected['value']

def ask_logo_effect_listbox(parent=None, initial='cut'):
    effects = [
        ('cut', 'Cut (đổi ngay)'),
        ('zoom', 'Zoom'),
        ('slide', 'Slide'),
        ('cube', 'Cube'),
        ('fly', 'Fly'),
        ('mix', 'Hỗn hợp ngẫu nhiên (Mix)')
    ]
    default_effect = initial if initial in ('cut', 'zoom', 'slide', 'cube', 'fly', 'mix') else 'cut'
    selected = {'value': None}

    dlg = tk.Toplevel(parent) if parent is not None else tk.Toplevel()
    dlg.title('Chọn hiệu ứng chuyển logo')
    try:
        if parent is not None:
            dlg.transient(parent)
        dlg.grab_set()
    except Exception:
        pass

    tk.Label(dlg, text='Chọn hiệu ứng:', font=('Segoe UI', 11, 'bold')).pack(padx=12, pady=(10, 6))
    lb = tk.Listbox(dlg, height=len(effects), width=34, exportselection=False, font=('Segoe UI', 10))
    for key, label in effects:
        lb.insert('end', f'{key} - {label}')
    lb.pack(padx=12, pady=6)

    idx_default = 0
    for idx, (key, _) in enumerate(effects):
        if key == default_effect:
            idx_default = idx
            break
    try:
        lb.selection_clear(0, 'end')
        lb.selection_set(idx_default)
        lb.activate(idx_default)
        lb.see(idx_default)
    except Exception:
        pass

    def on_ok(event=None):
        try:
            cur = lb.curselection()
            if cur:
                selected['value'] = effects[cur[0]][0]
            else:
                selected['value'] = default_effect
        except Exception:
            selected['value'] = default_effect
        try:
            dlg.destroy()
        except Exception:
            pass

    def on_cancel(event=None):
        selected['value'] = None
        try:
            dlg.destroy()
        except Exception:
            pass

    btns = tk.Frame(dlg)
    btns.pack(padx=12, pady=(6, 10))
    tk.Button(btns, text='OK', width=12, command=on_ok).pack(side='left', padx=6)
    tk.Button(btns, text='Hủy', width=12, command=on_cancel).pack(side='left', padx=6)

    lb.bind('<Double-Button-1>', on_ok)
    dlg.bind('<Return>', on_ok)
    dlg.bind('<Escape>', on_cancel)

    try:
        dlg.wait_window()
    except Exception:
        pass
    return selected['value']

def fetch_matches_from_sheet(sheet_url: str, max_rows: int = None):
    """Return list of row dicts from Google Sheet, or {'error': msg} on failure.

    - If URL contains `gid`, ưu tiên đọc đúng tab theo gid.
    - Nếu không có gid, fallback theo các tên tab phổ biến.
    Uses `fetch_matches_from_sheet._creds_path` for credentials.
    """

    def _normalize_text(value):
        import re
        import unicodedata
        s = '' if value is None else str(value)
        s = unicodedata.normalize('NFD', s)
        s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
        s = s.lower().replace('_', ' ').replace('-', ' ')
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    def _find_key(headers, *candidates):
        normalized = {_normalize_text(h): h for h in headers}
        for c in candidates:
            nc = _normalize_text(c)
            for nk, orig in normalized.items():
                if nk == nc:
                    return orig
        for c in candidates:
            nc = _normalize_text(c)
            for nk, orig in normalized.items():
                if nc and (nc in nk or nk in nc):
                    return orig
        return None

    def _pick_header_row(values):
        best_idx = None
        best_score = -1
        scan_limit = min(len(values), 200)
        for idx in range(scan_limit):
            row = values[idx] if idx < len(values) else []
            if not row:
                continue
            row_norm = [_normalize_text(x) for x in row]
            non_empty = sum(1 for x in row_norm if x)
            has_tran = any(x in ('tran', 'so tran', 'so tran dau') or 'tran' in x for x in row_norm)
            has_vdv_a = any(x in ('vdv a', 'vđv a', 'vdva') or 'vdv a' in x or 'vdva' in x for x in row_norm)
            has_vdv_b = any(x in ('vdv b', 'vđv b', 'vdvb') or 'vdv b' in x or 'vdvb' in x for x in row_norm)
            has_ban = any(x == 'ban' or x == 'bang' for x in row_norm)
            score = (3 if has_tran else 0) + (3 if has_vdv_a else 0) + (3 if has_vdv_b else 0) + (1 if has_ban else 0) + min(non_empty, 8)
            if score > best_score:
                best_score = score
                best_idx = idx
        return best_idx

    def _to_rows_from_values(values):
        if not values:
            return []

        header_idx = _pick_header_row(values)
        if header_idx is None:
            return []

        raw_headers = values[header_idx] if header_idx < len(values) else []
        headers = []
        seen = {}
        for i, h in enumerate(raw_headers):
            key = str(h).strip() if h is not None else ''
            if not key:
                key = f'COL_{i+1}'
            base = key
            n = 2
            while key in seen:
                key = f'{base}_{n}'
                n += 1
            seen[key] = True
            headers.append(key)

        rows = []
        for row in values[header_idx + 1:]:
            if not isinstance(row, list):
                continue
            obj = {}
            has_any = False
            for i, key in enumerate(headers):
                cell = row[i] if i < len(row) else ''
                sval = '' if cell is None else str(cell).strip()
                if sval:
                    has_any = True
                obj[key] = sval
            if has_any:
                rows.append(obj)
                if max_rows is not None:
                    try:
                        if len(rows) >= int(max_rows):
                            break
                    except Exception:
                        pass

        if not rows:
            return []

        tran_key = _find_key(headers, 'Trận', 'Số trận', 'tran', 'so tran')
        vdv_a_key = _find_key(headers, 'Tên VĐV A', 'VĐV A', 'VDV A', 'VĐVA', 'vdv a')
        vdv_b_key = _find_key(headers, 'Tên VĐV B', 'VĐV B', 'VDV B', 'VĐVB', 'vdv b')

        if not tran_key:
            return rows

        import re
        filtered = []
        for r in rows:
            tran_val = str(r.get(tran_key, '')).strip()
            if not tran_val:
                continue
            if not re.search(r'\d+', tran_val):
                continue
            if vdv_a_key or vdv_b_key:
                a = str(r.get(vdv_a_key, '')).strip() if vdv_a_key else ''
                b = str(r.get(vdv_b_key, '')).strip() if vdv_b_key else ''
                if not a and not b:
                    continue
            filtered.append(r)

        return filtered if filtered else rows

    try:
        if not sheet_url:
            return []

        import re
        m = re.search(r"/spreadsheets/d/([\w-]+)", sheet_url)
        spreadsheet_id = m.group(1) if m else None
        gid_match = re.search(r"[?#&]gid=(\d+)", sheet_url)
        gid = gid_match.group(1) if gid_match else None

        creds = getattr(fetch_matches_from_sheet, '_creds_path', None)
        if not (spreadsheet_id and creds and os.path.exists(creds)):
            return {'error': 'Missing spreadsheet ID or credentials (choose credentials)'}

        gs_client_cls = globals().get('GSheetClient', None)
        if gs_client_cls is None:
            try:
                from src.gsheet_client import GSheetClient as gs_client_cls
            except Exception as ie:
                return {'error': f'Cannot import GSheetClient: {ie}'}

        gs = gs_client_cls(spreadsheet_id, creds)

        meta = None
        try:
            meta = gs.get_metadata() if hasattr(gs, 'get_metadata') else None
        except Exception:
            meta = None

        gid_tab = None
        existing_titles = []
        if meta and isinstance(meta, dict):
            for sheet in meta.get('sheets', []):
                props = sheet.get('properties', {})
                title = props.get('title')
                sheet_id = props.get('sheetId')
                if title:
                    existing_titles.append(title)
                if gid and str(sheet_id) == str(gid) and title:
                    gid_tab = title

        tab_candidates = []
        if gid:
            if gid_tab:
                tab_candidates = [gid_tab]
            else:
                return {'error': f'Không tìm thấy tab đúng với gid={gid} trong Google Sheet.'}
        else:
            preferred_tabs = ['Kết quả', 'KET QUA', 'Ket qua', 'Ket Qua']
            for t in preferred_tabs:
                if t in existing_titles and t not in tab_candidates:
                    tab_candidates.append(t)
            for t in preferred_tabs:
                if t not in tab_candidates:
                    tab_candidates.append(t)
            for t in existing_titles:
                if t not in tab_candidates:
                    tab_candidates.append(t)

        last_error = None
        for tab_name in tab_candidates:
            try:
                rng = f'{tab_name}!A1:Z2000'
                raw_values = gs.batch_get(rng) if hasattr(gs, 'batch_get') else None
                rows = _to_rows_from_values(raw_values) if raw_values is not None else []
                if not rows:
                    rows = gs.read_table(rng) or []
                    if max_rows is not None:
                        try:
                            rows = rows[:int(max_rows)]
                        except Exception:
                            pass
                if rows:
                    return rows
            except Exception as ex:
                last_error = ex
                continue

        if last_error is not None:
            if gid and tab_candidates:
                return {'error': f"Không đọc được dữ liệu từ tab {tab_candidates[0]} (gid={gid}). Chi tiết: {last_error}"}
            return {'error': f"Không đọc được dữ liệu từ các tab đã thử: {tab_candidates}. Chi tiết: {last_error}"}
        return []
    except Exception as ex:
        return {'error': str(ex)}


def fetch_matches_from_ketqua(sheet_url: str, max_rows: int = None):
    return fetch_matches_from_sheet(sheet_url, max_rows)

class FullScreenMatchGUI(tk.Tk):
    def __init__(self):
        _enable_high_dpi_awareness()
        super().__init__()
        try:
            self._state_schema_version = 2
        except Exception:
            pass
        try:
            self._sync_tk_scaling_with_windows_dpi()
        except Exception:
            pass
        # ensure some commonly-used instance attributes exist before other init steps
        try:
            self.match_rows = []
            self.sheet_rows = []
            self.sheet_url = ''
            self.num_ban = None
            self.creds_path = None
            try:
                fallback_creds = [
                    os.path.join(os.getcwd(), 'credentials.json'),
                    os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'credentials.json')),
                ]
                for cp in fallback_creds:
                    if cp and os.path.exists(cp):
                        self.creds_path = cp
                        break
            except Exception:
                pass
        except Exception:
            pass
        
        try:
            self._auto_state_path = self._resolve_auto_state_path()
        except Exception:
            try:
                self._auto_state_path = os.path.join(os.getcwd(), 'ui_state.pkl')
            except Exception:
                self._auto_state_path = None
        try:
            # last-known preview snapshot (serializable) preserved even if preview window closed
            self._last_preview_meta = None
        except Exception:
            pass
        try:
            # True khi người dùng đã explicitly cấu hình qua dialog 'Cấu hình Preview'
            self._preview_meta_user_configured = False
        except Exception:
            pass
        try:
            self._preview_footer_logo_path = ''
        except Exception:
            pass
        try:
            # used to briefly suspend periodic autosave after restore to avoid overwrite races
            self._autosave_suspended_until = 0
        except Exception:
            pass
        try:
            # indicate whether a restore has been fully committed to disk
            # when False, autosave will skip writes until restore finishes
            self._restore_committed = True
        except Exception:
            pass
        try:
            # Auto-run interval for per-row "Kết quả" logic (configurable in code).
            self.auto_ketqua_interval_ms = 5000
            self._auto_ketqua_enabled = True
            self._auto_ketqua_running = False
        except Exception:
            pass
        try:
            self._send_blink_rows = set()
            self._send_blink_phase = False
            self._send_blink_job = None
        except Exception:
            pass
        try:
            self._set_ban_blink_rows = set()
            self._set_ban_blink_phase = False
            self._set_ban_blink_job = None
        except Exception:
            pass
        # Attempt to eagerly load saved state so restore runs during widget init
        try:
            import pickle
            for p in self._iter_state_candidates():
                try:
                    if p and os.path.exists(p):
                        with open(p, 'rb') as f:
                            raw_state = pickle.load(f)
                        s = self._normalize_state_dict(raw_state)
                        if not s:
                            continue
                        self._restored_state = s
                        try:
                            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                            import datetime
                            with open(dbg, 'a', encoding='utf-8') as df:
                                df.write(f"[{datetime.datetime.now().isoformat()}] EAGER_RESTORE loaded={p} keys={list(s.keys())} table_rows={len(s.get('table',[]))}\n")
                        except Exception:
                            pass
                        break
                except Exception:
                    continue
        except Exception:
            pass
        # ensure basic control variables exist before building widgets
        try:
            self.ban_var = tk.IntVar(value=8)
        except Exception:
            try:
                self.ban_var = None
            except Exception:
                pass
        try:
            # build UI widgets (create_widgets contains the table and bottom bar)
            self.create_widgets()
        except Exception as _ex:
            try:
                import sys, traceback
                print('ERROR: create_widgets() raised:', _ex, file=sys.stderr)
                traceback.print_exc()
            except Exception:
                pass
        # UI construction and table population occur later in __init__
        # start maximized (keep titlebar so user can minimize/close)
        try:
            self._is_fullscreen = False
            try:
                # prefer maximized window (keeps titlebar) instead of true fullscreen
                self.attributes('-fullscreen', False)
            except Exception:
                pass
            try:
                # set a clear window title
                self.title('Quản lý Trận đấu - vMix Scoreboard')
            except Exception:
                pass
            try:
                self.state('zoomed')
            except Exception:
                try:
                    # fallback to geometry maximizing to screen size
                    w, h = self.winfo_screenwidth(), self.winfo_screenheight()
                    self.geometry(f"{w}x{h}+0+0")
                except Exception:
                    pass
            # Add keybindings to toggle pseudo-fullscreen or quit
            try:
                self.bind('<F11>', lambda e: self._toggle_fullscreen())
                self.bind('<Escape>', lambda e: self._toggle_fullscreen())
                self.bind('<Control-q>', lambda e: self._on_close_and_save_state())
            except Exception:
                pass
        except Exception:
            pass
        # ensure graceful close and always autosave
        try:
            self.protocol('WM_DELETE_WINDOW', self._on_close_and_save_state)
        except Exception:
            pass
        try:
            # also register autosave on normal process exit as a fallback
            atexit.register(self._auto_save_state)
        except Exception:
            pass
        # Eagerly load state on startup
        try:
            self._auto_restore_state_to_ui()
        except Exception:
            pass
        # periodic autosave will be scheduled after widgets are built and restore attempted

    def _sync_tk_scaling_with_windows_dpi(self):
        """Normalize Tk scaling so frozen exe matches script font sizing on high-DPI Windows."""
        try:
            current = float(self.tk.call('tk', 'scaling'))
        except Exception:
            return

        target = current

        # Optional manual override for diagnostics/tuning.
        try:
            override = os.environ.get('QUANLIL_TK_SCALING', '').strip()
            if override:
                target = float(override)
        except Exception:
            override = ''

        if not override and os.name == 'nt':
            dpi = None
            try:
                import ctypes
                user32 = ctypes.windll.user32

                hwnd = 0
                try:
                    hwnd = int(self.winfo_id())
                except Exception:
                    hwnd = 0

                # Best source: effective DPI of this window.
                if hwnd and hasattr(user32, 'GetDpiForWindow'):
                    try:
                        dpi = int(user32.GetDpiForWindow(hwnd))
                    except Exception:
                        dpi = None

                # Fallback: process/system DPI.
                if (not dpi) and hasattr(user32, 'GetDpiForSystem'):
                    try:
                        dpi = int(user32.GetDpiForSystem())
                    except Exception:
                        dpi = None

                # Last fallback for old systems.
                if not dpi:
                    try:
                        hdc = user32.GetDC(0)
                        if hdc:
                            log_pixels_x = 88
                            gdi32 = ctypes.windll.gdi32
                            dpi = int(gdi32.GetDeviceCaps(hdc, log_pixels_x))
                            user32.ReleaseDC(0, hdc)
                    except Exception:
                        dpi = None
            except Exception:
                dpi = None

            # Tk scaling expects pixels per point (1 point = 1/72 inch).
            if dpi and dpi > 0:
                target = float(dpi) / 72.0

        try:
            if target > 0 and abs(target - current) >= 0.02:
                self.tk.call('tk', 'scaling', target)
        except Exception:
            pass

    def _resolve_auto_state_path(self):
        """Return the canonical ui_state.pkl path for this runtime mode."""
        try:
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(os.path.abspath(sys.executable))
                if exe_dir:
                    return os.path.join(exe_dir, 'ui_state.pkl')
        except Exception:
            pass

        try:
            project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
            return os.path.join(project_root, 'ui_state.pkl')
        except Exception:
            pass

        return os.path.join(os.getcwd(), 'ui_state.pkl')

    def _resolve_credentials_path(self, candidate=None):
        """Resolve a usable credentials file path on the current machine."""
        candidates = []

        exe_dir = ''
        try:
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        except Exception:
            exe_dir = ''

        try:
            project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        except Exception:
            project_root = os.getcwd()

        def _add_path(p):
            try:
                if p:
                    candidates.append(str(p))
            except Exception:
                pass

        if candidate:
            _add_path(candidate)
            try:
                base = os.path.basename(str(candidate))
            except Exception:
                base = ''

            try:
                if not os.path.isabs(str(candidate)):
                    if exe_dir:
                        _add_path(os.path.join(exe_dir, str(candidate)))
                    _add_path(os.path.join(os.getcwd(), str(candidate)))
                    _add_path(os.path.join(project_root, str(candidate)))
            except Exception:
                pass

            if base:
                if exe_dir:
                    _add_path(os.path.join(exe_dir, base))
                _add_path(os.path.join(os.getcwd(), base))
                _add_path(os.path.join(project_root, base))

        for fn in ('credentials.json', 'service_account.json'):
            if exe_dir:
                _add_path(os.path.join(exe_dir, fn))
            _add_path(os.path.join(os.getcwd(), fn))
            _add_path(os.path.join(project_root, fn))

        seen = set()
        for p in candidates:
            try:
                ap = os.path.abspath(p)
                key = os.path.normcase(ap)
            except Exception:
                ap = p
                key = p
            if key in seen:
                continue
            seen.add(key)
            try:
                if os.path.exists(ap):
                    return ap
            except Exception:
                continue
        return None

    def _friendly_network_error(self, ex, vmix_url=''):
        """Convert low-level request/connection errors to concise operator-friendly text."""
        try:
            raw = str(ex or '').strip()
        except Exception:
            raw = 'Lỗi mạng không xác định'

        host = ''
        try:
            from urllib.parse import urlparse
            if vmix_url:
                parsed = urlparse(vmix_url if '://' in vmix_url else f'http://{vmix_url}')
                host = parsed.hostname or ''
        except Exception:
            host = ''

        low = raw.lower()
        host_label = host or 'vMix host'

        if 'newconnectionerror' in low or 'failed to establish a new connection' in low:
            return f'Không kết nối được tới {host_label}. Kiểm tra mạng LAN, DNS/hosts hoặc dùng IP thay hostname.'
        if 'name or service not known' in low or 'nodename nor servname provided' in low or 'getaddrinfo failed' in low:
            return f'Không phân giải được tên máy {host_label}. Hãy đổi URL vMix sang địa chỉ IP.'
        if 'timed out' in low or 'read timed out' in low or 'connect timeout' in low:
            return f'Hết thời gian kết nối tới {host_label}. Kiểm tra vMix API và firewall cổng 8088.'
        if 'connection refused' in low:
            return f'{host_label} từ chối kết nối. Kiểm tra vMix đang chạy và API đã bật.'

        # Fallback: keep concise, avoid dumping full urllib trace to status bar.
        short = raw
        try:
            if len(short) > 180:
                short = short[:177] + '...'
        except Exception:
            pass
        return f'Lỗi mạng tới {host_label}: {short}'

    def _iter_state_candidates(self):
        """Yield primary state file first, then legacy fallback locations once."""
        seen = set()
        candidates = []
        primary = getattr(self, '_auto_state_path', None)
        if primary:
            candidates.append(primary)
        candidates.append(os.path.join(os.getcwd(), 'ui_state.pkl'))
        try:
            candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ui_state.pkl'))
        except Exception:
            pass
        try:
            candidates.append(os.path.join(os.path.dirname(os.path.abspath(sys.executable)), 'ui_state.pkl'))
        except Exception:
            pass

        for p in candidates:
            try:
                norm = os.path.normcase(os.path.abspath(p))
            except Exception:
                norm = str(p)
            if not p or norm in seen:
                continue
            seen.add(norm)
            yield p

    def _normalize_state_dict(self, state):
        """Normalize legacy state to current schema and drop stale preview payloads."""
        if not isinstance(state, dict):
            return None

        out = dict(state)
        try:
            schema = int(out.get('_schema_version', 0) or 0)
        except Exception:
            schema = 0

        # Legacy files (without schema) can carry outdated preview layout assumptions.
        # Keep business data but reset preview config so new renderer starts clean.
        if schema < 2:
            out['preview'] = None
            out['preview_footer_logo'] = str(out.get('preview_footer_logo') or '')

        # Credentials path from another machine is often invalid in portable usage.
        try:
            cred = out.get('creds_path')
            if cred:
                resolved = self._resolve_credentials_path(cred)
                out['creds_path'] = resolved
        except Exception:
            pass

        # Mark if state seems to be generated on another machine (used for soft warning only).
        try:
            import platform
            current_node = str(platform.node() or '').strip().lower()
            saved_node = str(out.get('_machine_name') or '').strip().lower()
            out['_state_foreign_machine'] = bool(saved_node and current_node and saved_node != current_node)
        except Exception:
            out['_state_foreign_machine'] = False

        out['_schema_version'] = int(getattr(self, '_state_schema_version', 2) or 2)
        return out

    def fetch_all_vmix_to_table(self):
        """Fetch vMix Input 1 for every row's vMix URL and populate table fields.
        This updates the in-window entries (Tên A, Tên B, Điểm) where possible.
        """
        try:
            import requests, xml.etree.ElementTree as ET, time
        except Exception:
            return 0
        updated = 0
        for idx, widgets in enumerate(getattr(self, 'match_rows', []) or []):
            try:
                vmix_url = widgets[5].get().strip() if len(widgets) > 5 and hasattr(widgets[5], 'get') else ''
            except Exception:
                vmix_url = ''
            if not vmix_url:
                continue
            try:
                resp = requests.get(f"{vmix_url}/API/", timeout=2)
                root = ET.fromstring(resp.text)
                input1 = root.find(".//input[@number='1']")
                if input1 is None:
                    continue
                def get_field(name):
                    for t in input1.findall('text'):
                        if t.attrib.get('name') == name:
                            return t.text or ''
                    return ''
                tenA = get_field('TenA') or get_field('TenA.Text')
                tenB = get_field('TenB') or get_field('TenB.Text')
                diemA = get_field('DiemA') or get_field('DiemA.Text')
                diemB = get_field('DiemB') or get_field('DiemB.Text')
                # update widgets (2: Tên A, 3: Tên B, 4: Điểm)
                try:
                    if len(widgets) > 2 and hasattr(widgets[2], 'config'):
                        widgets[2].config(state='normal')
                        widgets[2].delete(0, 'end')
                        widgets[2].insert(0, tenA)
                        widgets[2].config(state='readonly')
                except Exception:
                    pass
                try:
                    if len(widgets) > 3 and hasattr(widgets[3], 'config'):
                        widgets[3].config(state='normal')
                        widgets[3].delete(0, 'end')
                        widgets[3].insert(0, tenB)
                        widgets[3].config(state='readonly')
                except Exception:
                    pass
                try:
                    if len(widgets) > 4 and hasattr(widgets[4], 'delete'):
                        # main table has a single 'Điểm' cell; set as 'A-B' if both present
                        val = ''
                        if diemA and diemB:
                            val = f"{diemA}-{diemB}"
                        elif diemA:
                            val = str(diemA)
                        elif diemB:
                            val = str(diemB)
                        widgets[4].delete(0, 'end')
                        widgets[4].insert(0, val)
                except Exception:
                    pass
                updated += 1
                # small pause to avoid hammering many vMix hosts
                try:
                    time.sleep(0.05)
                except Exception:
                    pass
            except Exception:
                # continue on any per-row error
                continue
        try:
            self.status_var.set(f'Đã cập nhật {updated} dòng từ vMix')
        except Exception:
            pass
        return updated

    def fetch_all_vmix_to_table_async(self):
        """Run fetch_all_vmix_to_table in a background thread and update UI when done."""
        try:
            import threading
        except Exception:
            threading = None
        if getattr(self, '_fetch_thread_running', False):
            try:
                self.status_var.set('Đang cập nhật vMix... (vẫn chạy)')
            except Exception:
                pass
            return
        def _worker():
            try:
                self._fetch_thread_running = True
                try:
                    count = self.fetch_all_vmix_to_table()
                except Exception:
                    count = 0
                try:
                    self.after(10, lambda: self.status_var.set(f'Đã cập nhật {count} dòng từ vMix'))
                except Exception:
                    pass
            finally:
                try:
                    self._fetch_thread_running = False
                except Exception:
                    pass
        try:
            if threading:
                t = threading.Thread(target=_worker, daemon=True)
                t.start()
            else:
                _worker()
        except Exception:
            try:
                _worker()
            except Exception:
                pass

    def write_all_vmix_to_sheet(self, range_read='Kết quả!A1:Z2000'):
        """Batch-write vMix fields for all rows into Google Sheets.

        For each GUI row that has a vMix URL and a Trận value, this will fetch
        Input 1 from vMix, map common fields (Tên, Điểm, Lco, HRs, AVGs) to the
        sheet headers and perform a safe batch update (only writes AA..AL by default).
        Returns the number of rows updated or -1 on error.
        """
        try:
            import requests, xml.etree.ElementTree as ET, os
        except Exception:
            return -1
        # ensure sheets client available
        sheet_url = self.url_var.get().strip() if hasattr(self, 'url_var') else ''
        m = None
        if sheet_url:
            try:
                import re
                m = re.search(r"/spreadsheets/d/([\w-]+)", sheet_url)
            except Exception:
                m = None
        spreadsheet_id = m.group(1) if m else None
        creds = getattr(self, 'creds_path', None)
        if not (spreadsheet_id and creds and os.path.exists(creds)):
            try:
                self.status_var.set('Chưa cấu hình Google Sheets (URL/credentials)')
            except Exception:
                pass
            return -1
        try:
            gs = GSheetClient(spreadsheet_id, creds)
        except Exception as ex:
            try:
                self.status_var.set(f'Không thể khởi tạo GSheetClient: {ex}')
            except Exception:
                pass
            return -1

        # read existing sheet to locate rows and headers
        try:
            rows = gs.read_table(range_read)
            headers = list(rows[0].keys()) if rows else []
        except Exception:
            rows = []
            headers = []

        def find_col_key(keys, *candidates):
            def normalize(s):
                return s.replace(' ', '').replace('_', '').lower()
            norm_keys = {normalize(k): k for k in keys}
            for c in candidates:
                nc = normalize(c)
                if nc in norm_keys:
                    return norm_keys[nc]
            # fallback to letter index
            for c in candidates:
                if len(c) == 1 and c.isalpha():
                    idx = ord(c.upper()) - ord('A')
                    if idx < len(keys):
                        return keys[idx]
            return None

        batch = []
        updates = 0
        for idx, widgets in enumerate(getattr(self, 'match_rows', []) or []):
            try:
                tran_val = widgets[0].get().strip() if hasattr(widgets[0], 'get') else ''
                vmix_url = widgets[5].get().strip() if len(widgets) > 5 and hasattr(widgets[5], 'get') else ''
            except Exception:
                continue
            if not tran_val or not vmix_url:
                continue
            # find sheet row by matching 'Trận' column
            found_idx = None
            for ridx, r in enumerate(rows):
                try:
                    # attempt to compare numeric tokens
                    import re
                    m1 = re.search(r"(\d+)", str(tran_val))
                    m2 = re.search(r"(\d+)", str(r.get(find_col_key(headers, 'Trận') if headers else ''))) 
                    if m1 and m2 and int(m1.group(1)) == int(m2.group(1)):
                        found_idx = ridx
                        break
                except Exception:
                    continue
            if found_idx is None:
                continue

            # fetch vMix fields
            try:
                resp = requests.get(f"{vmix_url}/API/", timeout=2)
                root = ET.fromstring(resp.text)
                input1 = root.find(".//input[@number='1']")
                def get_field(name):
                    if input1 is None:
                        return ''
                    for t in input1.findall('text'):
                        if t.attrib.get('name') == name:
                            return t.text or ''
                    return ''
                # map fields
                diem_a = get_field('DiemA') or get_field('DiemA.Text')
                diem_b = get_field('DiemB') or get_field('DiemB.Text')
                lco = get_field('Lco') or get_field('Lco.Text')
                ten_a = get_field('TenA') or get_field('TenA.Text')
                ten_b = get_field('TenB') or get_field('TenB.Text')
                hr1a = get_field('HR1A') or get_field('HR1A.Text')
                hr2a = get_field('HR2A') or get_field('HR2A.Text')
                hr1b = get_field('HR1B') or get_field('HR1B.Text')
                hr2b = get_field('HR2B') or get_field('HR2B.Text')
                # compute AVGs
                def to_float(v):
                    try:
                        return float(str(v).replace(',', '.'))
                    except Exception:
                        return None
                a = to_float(diem_a); b = to_float(diem_b); c = to_float(lco)
                avga = (round(a/c,3) if a is not None and c and c != 0 else '')
                avgb = (round(b/c,3) if b is not None and c and c != 0 else '')

                # find header keys
                ten_a_col = find_col_key(headers, 'Tên VĐV A', 'TenA', 'TenA.Text')
                ten_b_col = find_col_key(headers, 'Tên VĐV B', 'TenB', 'TenB.Text')
                avga_col = find_col_key(headers, 'AVGA', 'AvgA', 'AvgA.Text')
                avgb_col = find_col_key(headers, 'AVGB', 'AvgB', 'AvgB.Text')
                diem_a_col = find_col_key(headers, 'Điểm A', 'DiemA', 'DiemA.Text')
                diem_b_col = find_col_key(headers, 'Điểm B', 'DiemB', 'DiemB.Text')
                lco_col = find_col_key(headers, 'Lượt cơ', 'Lco', 'Lco.Text')
                hr1a_col = find_col_key(headers, 'HR1A', 'HR1A.Text')
                hr2a_col = find_col_key(headers, 'HR2A', 'HR2A.Text')
                hr1b_col = find_col_key(headers, 'HR1B', 'HR1B.Text')
                hr2b_col = find_col_key(headers, 'HR2B', 'HR2B.Text')

                vals_map = {
                    ten_a_col: ten_a,
                    ten_b_col: ten_b,
                    diem_a_col: diem_a,
                    diem_b_col: diem_b,
                    lco_col: lco,
                    hr1a_col: hr1a,
                    hr2a_col: hr2a,
                    hr1b_col: hr1b,
                    hr2b_col: hr2b,
                    avga_col: (f"{avga:.3f}".replace('.', ',') if avga != '' else ''),
                    avgb_col: (f"{avgb:.3f}".replace('.', ',') if avgb != '' else '')
                }

                rownum = found_idx + 2
                for header_key, value in vals_map.items():
                    if not header_key:
                        continue
                    try:
                        col_idx = headers.index(header_key)
                    except ValueError:
                        continue
                    # restrict writes to AA..AL (columns 27..38)
                    try:
                        col_number = col_idx + 1
                        if col_number < 27 or col_number > 38:
                            continue
                    except Exception:
                        pass
                    col_letter = self.index_to_col_letter(col_idx)
                    cell_range = f'Kết quả!{col_letter}{rownum}'
                    batch.append({'range': cell_range, 'values': [[value]]})
                updates += 1
            except Exception:
                continue

        # execute batch updates in safe mode
        try:
            if batch:
                res = gs.batch_update_safe(batch)
                if not res:
                    try:
                        gs.batch_update(batch)
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            self.status_var.set(f'Đã ghi {updates} dòng lên Google Sheet')
        except Exception:
            pass
        return updates

    def preview_write_all_vmix_to_sheet(self, range_read='Kết quả!A1:Z2000'):
        """Prepare the batch that would be written by write_all_vmix_to_sheet but do not perform any network writes.
        Returns a tuple (batch_list, updates_count, summary_dict).
        """
        try:
            import requests, xml.etree.ElementTree as ET, os
        except Exception:
            return [], 0, {'error': 'missing deps'}
        # similar header/read logic but skip GSheet client
        sheet_url = self.url_var.get().strip() if hasattr(self, 'url_var') else ''
        m = None
        if sheet_url:
            try:
                import re
                m = re.search(r"/spreadsheets/d/([\w-]+)", sheet_url)
            except Exception:
                m = None
        spreadsheet_id = m.group(1) if m else None
        creds = getattr(self, 'creds_path', None)
        try:
            # attempt to read headers if possible, else empty
            headers = []
            if spreadsheet_id and creds and os.path.exists(creds):
                try:
                    gs = GSheetClient(spreadsheet_id, creds)
                    rows = gs.read_table(range_read)
                    headers = list(rows[0].keys()) if rows else []
                except Exception:
                    headers = []
        except Exception:
            headers = []

        def find_col_key(keys, *candidates):
            def normalize(s):
                return s.replace(' ', '').replace('_', '').lower()
            norm_keys = {normalize(k): k for k in keys}
            for c in candidates:
                nc = normalize(c)
                if nc in norm_keys:
                    return norm_keys[nc]
            for c in candidates:
                if len(c) == 1 and c.isalpha():
                    idx = ord(c.upper()) - ord('A')
                    if idx < len(keys):
                        return keys[idx]
            return None

        batch = []
        updates = 0
        for idx, widgets in enumerate(getattr(self, 'match_rows', []) or []):
            try:
                tran_val = widgets[0].get().strip() if hasattr(widgets[0], 'get') else ''
                vmix_url = widgets[5].get().strip() if len(widgets) > 5 and hasattr(widgets[5], 'get') else ''
            except Exception:
                continue
            if not tran_val or not vmix_url:
                continue
            # find sheet row index by matching 'Trận' header if headers known
            found_idx = None
            if headers:
                for ridx, r in enumerate(rows):
                    try:
                        import re
                        m1 = re.search(r"(\d+)", str(tran_val))
                        tran_col = find_col_key(headers, 'Trận')
                        sheet_val = r.get(tran_col, '') if tran_col else ''
                        m2 = re.search(r"(\d+)", str(sheet_val))
                        if m1 and m2 and int(m1.group(1)) == int(m2.group(1)):
                            found_idx = ridx
                            break
                    except Exception:
                        continue
            if found_idx is None:
                # cannot map row without sheet data; skip in preview
                continue
            # fetch vMix input1
            try:
                resp = requests.get(f"{vmix_url}/API/", timeout=2)
                root = ET.fromstring(resp.text)
                input1 = root.find(".//input[@number='1']")
                def get_field(name):
                    if input1 is None:
                        return ''
                    for t in input1.findall('text'):
                        if t.attrib.get('name') == name:
                            return t.text or ''
                    return ''
                diem_a = get_field('DiemA') or get_field('DiemA.Text')
                diem_b = get_field('DiemB') or get_field('DiemB.Text')
                lco = get_field('Lco') or get_field('Lco.Text')
                ten_a = get_field('TenA') or get_field('TenA.Text')
                ten_b = get_field('TenB') or get_field('TenB.Text')
                # compute avgs
                def to_float(v):
                    try:
                        return float(str(v).replace(',', '.'))
                    except Exception:
                        return None
                a = to_float(diem_a); b = to_float(diem_b); c = to_float(lco)
                avga = (round(a/c,3) if a is not None and c and c != 0 else '')
                avgb = (round(b/c,3) if b is not None and c and c != 0 else '')
                # find header keys
                ten_a_col = find_col_key(headers, 'Tên VĐV A', 'TenA', 'TenA.Text')
                ten_b_col = find_col_key(headers, 'Tên VĐV B', 'TenB', 'TenB.Text')
                avga_col = find_col_key(headers, 'AVGA', 'AvgA', 'AvgA.Text')
                avgb_col = find_col_key(headers, 'AVGB', 'AvgB', 'AvgB.Text')
                diem_a_col = find_col_key(headers, 'Điểm A', 'DiemA', 'DiemA.Text')
                diem_b_col = find_col_key(headers, 'Điểm B', 'DiemB', 'DiemB.Text')
                lco_col = find_col_key(headers, 'Lượt cơ', 'Lco', 'Lco.Text')
                hr1a_col = find_col_key(headers, 'HR1A', 'HR1A.Text')
                hr2a_col = find_col_key(headers, 'HR2A', 'HR2A.Text')
                hr1b_col = find_col_key(headers, 'HR1B', 'HR1B.Text')
                hr2b_col = find_col_key(headers, 'HR2B', 'HR2B.Text')

                vals_map = {
                    ten_a_col: ten_a,
                    ten_b_col: ten_b,
                    diem_a_col: diem_a,
                    diem_b_col: diem_b,
                    lco_col: lco,
                    hr1a_col: get_field('HR1A') or get_field('HR1A.Text'),
                    hr2a_col: get_field('HR2A') or get_field('HR2A.Text'),
                    hr1b_col: get_field('HR1B') or get_field('HR1B.Text'),
                    hr2b_col: get_field('HR2B') or get_field('HR2B.Text'),
                    avga_col: (f"{avga:.3f}".replace('.', ',') if avga != '' else ''),
                    avgb_col: (f"{avgb:.3f}".replace('.', ',') if avgb != '' else '')
                }
                # build batch ranges
                for header_key, value in vals_map.items():
                    if not header_key:
                        continue
                    try:
                        col_idx = headers.index(header_key)
                    except Exception:
                        continue
                    col_letter = self.index_to_col_letter(col_idx)
                    cell_range = f'Kết quả!{col_letter}{found_idx+2}'
                    batch.append({'range': cell_range, 'values': [[value]]})
                updates += 1
            except Exception:
                continue

        summary = {'spreadsheet_id': spreadsheet_id, 'creds_present': bool(creds and os.path.exists(creds)), 'updates': updates, 'batch_len': len(batch)}
        try:
            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
            import datetime
            with open(dbg, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now().isoformat()}] PREVIEW_WRITE summary={summary}\n")
        except Exception:
            pass
        return batch, updates, summary

    def write_all_vmix_to_sheet_async(self):
        """Run write_all_vmix_to_sheet in a background thread to avoid blocking UI."""
        try:
            import threading
        except Exception:
            threading = None
        if getattr(self, '_write_thread_running', False):
            try:
                self.status_var.set('Đang ghi lên Sheets... (vẫn chạy)')
            except Exception:
                pass
            return
        def _worker():
            try:
                self._write_thread_running = True
                try:
                    count = self.write_all_vmix_to_sheet()
                except Exception:
                    count = -1
                try:
                    self.after(10, lambda: self.status_var.set(f'Ghi xong: {count} dòng'))
                except Exception:
                    pass
            finally:
                try:
                    self._write_thread_running = False
                except Exception:
                    pass
        try:
            if threading:
                t = threading.Thread(target=_worker, daemon=True)
                t.start()
            else:
                _worker()
        except Exception:
            try:
                _worker()
            except Exception:
                pass

    def show_preview_write_popup(self):
        """Open a popup showing the dry-run batch that would be written to Google Sheets."""
        try:
            batch, updates, summary = self.preview_write_all_vmix_to_sheet()
        except Exception as ex:
            batch, updates, summary = [], 0, {'error': str(ex)}
        try:
            from tkinter import ttk
            dlg = tk.Toplevel(self)
            dlg.title('Preview: Ghi tất cả lên Sheets')
            dlg.geometry('900x600')
            frame = tk.Frame(dlg)
            frame.pack(fill='both', expand=True)
            cols = ('range', 'values', 'diff')
            tree = ttk.Treeview(frame, columns=cols, show='headings')
            tree.heading('range', text='Range')
            tree.heading('values', text='Values')
            tree.heading('diff', text='Diff')
            tree.column('range', width=220, anchor='w')
            tree.column('values', width=550, anchor='w')
            tree.column('diff', width=120, anchor='center')
            vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
            hsb = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
            tree.configure(yscroll=vsb.set, xscroll=hsb.set)
            tree.grid(row=0, column=0, sticky='nsew')
            vsb.grid(row=0, column=1, sticky='ns')
            hsb.grid(row=1, column=0, sticky='ew')
            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)
            # attempt to compute diffs if we have credentials and spreadsheet id
            can_diff = False
            spreadsheet_id = summary.get('spreadsheet_id') or getattr(self, 'spreadsheet_id', None) or getattr(self, 'sheet_url', None)
            creds_path = getattr(self, 'creds_path', None)
            gs = None
            try:
                if spreadsheet_id and creds_path and os.path.exists(creds_path):
                    import sys
                    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
                    from src.gsheet_client import GSheetClient
                    gs = GSheetClient(spreadsheet_id, creds_path)
                    can_diff = True
            except Exception:
                gs = None
                can_diff = False

            for i, it in enumerate(batch):
                rng = it.get('range') or f'item_{i}'
                vals = it.get('values')
                try:
                    vals_repr = ', '.join([str(x) for row in vals for x in row]) if vals else ''
                except Exception:
                    vals_repr = repr(vals)
                diff_flag = 'N/A'
                changed = False
                if can_diff and rng:
                    try:
                        current = gs.batch_get(rng)
                        # normalize empty vs missing
                        cur = current if current is not None else []
                        new = vals if vals is not None else []
                        if cur != new:
                            changed = True
                            diff_flag = 'CHANGED'
                        else:
                            diff_flag = 'UNCHANGED'
                    except Exception:
                        diff_flag = 'Err'
                tree.insert('', 'end', iid=str(i), values=(rng, vals_repr, diff_flag), tags=('changed',) if changed else ())
            # configure tag visuals
            try:
                tree.tag_configure('changed', background='#fff3e0')
            except Exception:
                pass
            info = tk.Label(dlg, text=f"Summary: {summary}    Batch items: {len(batch)}")
            info.pack(fill='x')
            btn_frame = tk.Frame(dlg)
            btn_frame.pack(fill='x')
            def select_all():
                for iid in tree.get_children():
                    tree.selection_add(iid)
            def deselect_all():
                for iid in tree.get_children():
                    tree.selection_remove(iid)
            tk.Button(btn_frame, text='Chọn tất cả', command=select_all).pack(side='left', padx=4, pady=6)
            tk.Button(btn_frame, text='Bỏ chọn', command=deselect_all).pack(side='left', padx=4, pady=6)
            def do_write_all():
                try:
                    dlg.destroy()
                except Exception:
                    pass
                try:
                    self.write_all_vmix_to_sheet_async()
                except Exception:
                    self.write_all_vmix_to_sheet()
            tk.Button(btn_frame, text='Ghi tất cả', bg='#00C853', fg='white', command=do_write_all).pack(side='right', padx=6, pady=6)
            tk.Button(btn_frame, text='Đóng', command=dlg.destroy).pack(side='right', padx=6, pady=6)
        except Exception as ex:
            try:
                messagebox.showerror('Lỗi', f'Không tạo được preview: {ex}')
            except Exception:
                pass

    def select_credentials(self):
        """Cho phép người dùng chọn file credentials.json và lưu đường dẫn vào `self.creds_path`."""
        try:
            from tkinter import filedialog, messagebox
            # import os đã ở đầu file
            path = filedialog.askopenfilename(title='Chọn credentials.json', filetypes=[('JSON files', '*.json'), ('All files', '*.*')])
            if path:
                self.creds_path = path
                try:
                    if hasattr(self, 'creds_label') and self.creds_label:
                        self.creds_label.config(text=os.path.basename(path))
                except Exception:
                    pass
        except Exception as ex:
            try:
                from tkinter import messagebox
                messagebox.showerror('Lỗi', f'Không chọn được credentials: {ex}')
            except Exception:
                pass

    # --- GHI BỔ SUNG: Ghi các trường vMix vào cột AA..AL (nếu ô rỗng) ---
    def index_to_col_letter(self, zero_based):
        n = zero_based + 1
        letters = ''
        while n > 0:
            n, rem = divmod(n-1, 26)
            letters = chr(65 + rem) + letters
        return letters

    def open_edit_popup(self, idx):
        """Open edit popup; prefill from live vMix Input 1 or CSV fallback."""
        import tkinter as tk
        from tkinter import messagebox
        import re, os, csv, sys
        try:
            # prepare sheet rows (optional)
            sheet_url = self.url_var.get().strip() if hasattr(self, 'url_var') else ''
            m = re.search(r"/spreadsheets/d/([\w-]+)", sheet_url) if sheet_url else None
            spreadsheet_id = m.group(1) if m else None
            creds_path = self.creds_path if hasattr(self, 'creds_path') else None
            gs = None
            rows = []
            if spreadsheet_id and creds_path and os.path.exists(creds_path):
                try:
                    gs = GSheetClient(spreadsheet_id, creds_path)
                    rows = gs.read_table('Kết quả!A1:Z2000') or []
                except Exception:
                    rows = []

            # build popup (dark themed, two-column A/B layout)
            popup = tk.Toplevel(self)
            popup.title('Sửa kết quả')
            popup.configure(bg='#162033')
            try:
                screen_w = popup.winfo_screenwidth()
                screen_h = popup.winfo_screenheight()
                pw = min(920, max(760, int(screen_w * 0.62)))
                ph = min(980, max(640, int(screen_h * 0.90)))
            except Exception:
                pw, ph = 780, 700
            popup.geometry(f'{pw}x{ph}')
            try:
                popup.minsize(700, 620)
                popup.resizable(True, True)
            except Exception:
                pass
            # keep transient but do not grab focus so main window remains visible
            try:
                popup.transient(self)
            except Exception:
                pass
            # center over parent window without covering content fully
            try:
                self.update_idletasks()
                px = self.winfo_rootx()
                py = self.winfo_rooty()
                pwid = self.winfo_width()
                phei = self.winfo_height()
                x = px + max(0, (pwid - pw) // 2)
                y = py + max(0, (phei - ph) // 2)
                max_x = max(0, popup.winfo_screenwidth() - pw)
                max_y = max(0, popup.winfo_screenheight() - ph)
                x = min(max(0, x), max_x)
                y = min(max(0, y), max_y)
                popup.geometry(f'{pw}x{ph}+{x}+{y}')
            except Exception:
                pass

            # top area: scrollable form so content is never truncated
            content_wrap = tk.Frame(popup, bg='#162033')
            content_wrap.pack(side='top', fill='both', expand=True)
            content_wrap.grid_rowconfigure(0, weight=1)
            content_wrap.grid_columnconfigure(0, weight=1)

            form_canvas = tk.Canvas(content_wrap, bg='#162033', highlightthickness=0, bd=0)
            form_scroll = tk.Scrollbar(content_wrap, orient='vertical', command=form_canvas.yview)
            form_canvas.configure(yscrollcommand=form_scroll.set)

            form_canvas.grid(row=0, column=0, sticky='nsew', padx=(8, 0), pady=(10, 0))
            form_scroll.grid(row=0, column=1, sticky='ns', padx=(0, 8), pady=(10, 0))

            form_bg = '#1F2A40'
            form_frame = tk.Frame(form_canvas, bg=form_bg, highlightbackground='#2E3D59', highlightthickness=1)
            form_window = form_canvas.create_window((0, 0), window=form_frame, anchor='nw')

            # Make form_frame expand with canvas
            def _resize_form(event):
                try:
                    form_canvas.itemconfigure(form_window, width=event.width)
                except Exception:
                    pass
            form_canvas.bind('<Configure>', _resize_form)

            def _sync_scrollregion(event=None):
                try:
                    form_canvas.configure(scrollregion=form_canvas.bbox('all'))
                except Exception:
                    pass

            def _sync_form_width(event=None):
                try:
                    if event is not None:
                        form_canvas.itemconfigure(form_window, width=event.width)
                except Exception:
                    pass

            form_frame.bind('<Configure>', _sync_scrollregion)
            form_canvas.bind('<Configure>', _sync_form_width)

            def _on_mousewheel(event):
                try:
                    delta = getattr(event, 'delta', 0)
                    if delta:
                        step = -1 if delta > 0 else 1
                        form_canvas.yview_scroll(step, 'units')
                    else:
                        num = getattr(event, 'num', None)
                        if num == 4:
                            form_canvas.yview_scroll(-1, 'units')
                        elif num == 5:
                            form_canvas.yview_scroll(1, 'units')
                except Exception:
                    pass

            for wheel_widget in (popup, form_canvas, form_frame):
                try:
                    wheel_widget.bind('<MouseWheel>', _on_mousewheel)
                    wheel_widget.bind('<Button-4>', _on_mousewheel)
                    wheel_widget.bind('<Button-5>', _on_mousewheel)
                except Exception:
                    pass

            entries = {}
            lbl_fg = '#FFD369'
            entry_bg = '#25344D'
            entry_fg = '#F5F7FA'
            fnt_label = ('Segoe UI', 14, 'bold')
            fnt_entry = ('Segoe UI', 14)
            entry_style = {
                'bg': entry_bg,
                'fg': entry_fg,
                'insertbackground': entry_fg,
                'font': fnt_entry,
                'relief': 'flat',
                'bd': 0,
                'highlightthickness': 1,
                'highlightbackground': '#3A4B69',
                'highlightcolor': '#5B8CFF',
            }

            # Column headings (A / B)
            tk.Label(form_frame, text='VĐV A', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=0, column=0, padx=12, pady=(12,6), sticky='w')
            tk.Label(form_frame, text='VĐV B', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=0, column=1, padx=12, pady=(12,6), sticky='w')

            # Tên
            tk.Label(form_frame, text='Tên', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=1, column=0, sticky='w', padx=12)
            e_tena = tk.Entry(form_frame, **entry_style)
            e_tena.grid(row=2, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Tên VĐV A'] = e_tena

            tk.Label(form_frame, text='Tên', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=1, column=1, sticky='w', padx=12)
            e_tenb = tk.Entry(form_frame, **entry_style)
            e_tenb.grid(row=2, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Tên VĐV B'] = e_tenb

            # Điểm row
            tk.Label(form_frame, text='Điểm', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=3, column=0, sticky='w', padx=12)
            e_diem_a = tk.Entry(form_frame, **entry_style)
            e_diem_a.grid(row=4, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Điểm A'] = e_diem_a

            tk.Label(form_frame, text='Điểm', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=3, column=1, sticky='w', padx=12)
            e_diem_b = tk.Entry(form_frame, **entry_style)
            e_diem_b.grid(row=4, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Điểm B'] = e_diem_b

            # Lượt cơ should appear under the Điểm row, spanning both columns
            tk.Label(form_frame, text='Lượt cơ', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=5, column=0, columnspan=2, sticky='w', padx=12)
            e_lco = tk.Entry(form_frame, **entry_style)
            e_lco.grid(row=6, column=0, columnspan=2, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Lượt cơ'] = e_lco

            # HR rows: HR1 and HR2 for A and B
            tk.Label(form_frame, text='HR1', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=7, column=0, sticky='w', padx=12)
            e_hr1a = tk.Entry(form_frame, **entry_style)
            e_hr1a.grid(row=8, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR1A'] = e_hr1a

            tk.Label(form_frame, text='HR1', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=7, column=1, sticky='w', padx=12)
            e_hr1b = tk.Entry(form_frame, **entry_style)
            e_hr1b.grid(row=8, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR1B'] = e_hr1b

            tk.Label(form_frame, text='HR2', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=9, column=0, sticky='w', padx=12)
            e_hr2a = tk.Entry(form_frame, **entry_style)
            e_hr2a.grid(row=10, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR2A'] = e_hr2a

            tk.Label(form_frame, text='HR2', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=9, column=1, sticky='w', padx=12)
            e_hr2b = tk.Entry(form_frame, **entry_style)
            e_hr2b.grid(row=10, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR2B'] = e_hr2b

            # AVGs
            tk.Label(form_frame, text='AVGA', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=11, column=0, sticky='w', padx=12)
            e_avga = tk.Entry(form_frame, **entry_style, state='normal')
            e_avga.grid(row=12, column=0, padx=12, pady=(0,12), sticky='ew', ipady=6)
            entries['AVGA'] = e_avga

            tk.Label(form_frame, text='AVGB', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=11, column=1, sticky='w', padx=12)
            e_avgb = tk.Entry(form_frame, **entry_style, state='normal')
            e_avgb.grid(row=12, column=1, padx=12, pady=(0,12), sticky='ew', ipady=6)
            entries['AVGB'] = e_avgb

            # Trận (đặt dưới cùng theo yêu cầu)
            tk.Label(form_frame, text='Trận', fg=lbl_fg, bg=form_bg, font=fnt_label).grid(row=13, column=0, columnspan=2, sticky='w', padx=12)
            e_tran = tk.Entry(form_frame, **entry_style)
            e_tran.grid(row=14, column=0, columnspan=2, padx=12, pady=(0,12), sticky='ew', ipady=6)
            entries['Trận'] = e_tran


            # Nếu có bàn hiện tại thì chọn sẵn
            try:
                if hasattr(self, 'match_rows') and idx < len(self.match_rows):
                    current_tran = self.match_rows[idx][0].get().strip()
                    if current_tran:
                        entries['Trận'].delete(0, 'end')
                        entries['Trận'].insert(0, current_tran)
            except Exception:
                pass

            # make grid expand entries
            form_frame.grid_columnconfigure(0, weight=1)
            form_frame.grid_columnconfigure(1, weight=1)

            def compute_avgs():
                try:
                    a = float(entries['Điểm A'].get()) if entries['Điểm A'].get() not in (None,'') else None
                except Exception:
                    a = None
                try:
                    b = float(entries['Điểm B'].get()) if entries['Điểm B'].get() not in (None,'') else None
                except Exception:
                    b = None
                try:
                    c = float(entries['Lượt cơ'].get()) if entries['Lượt cơ'].get() not in (None,'') else None
                except Exception:
                    c = None
                if a is not None and c and c != 0:
                    entries['AVGA'].delete(0,'end')
                    entries['AVGA'].insert(0, str(round(a/c,3)))
                if b is not None and c and c != 0:
                    entries['AVGB'].delete(0,'end')
                    entries['AVGB'].insert(0, str(round(b/c,3)))

            def prefill_from_vmix():
                vmix_url = ''
                try:
                    if hasattr(self, 'match_rows') and idx < len(self.match_rows):
                        vmix_url = self.match_rows[idx][5].get().strip()
                except Exception:
                    vmix_url = ''
                vmr = None
                if vmix_url:
                    try:
                        import requests
                        import xml.etree.ElementTree as ET
                        resp = requests.get(f"{vmix_url}/API/", timeout=2)
                        root = ET.fromstring(resp.text)
                        input1 = root.find(".//input[@number='1']")
                        if input1 is not None:
                            vmr = {t.attrib.get('name'): (t.text or '') for t in input1.findall('text')}
                    except Exception:
                        vmr = None
                if not vmr:
                    # fallback CSV
                    base = os.path.dirname(sys.executable) if hasattr(sys, '_MEIPASS') else os.path.dirname(os.path.abspath(__file__))
                    candidates = [os.path.join(base, 'vmix_input1_temp.csv'), os.path.join(os.getcwd(), 'vmix_input1_temp.csv')]
                    path = None
                    for p in candidates:
                        if os.path.exists(p):
                            path = p
                            break
                    if path:
                        try:
                            with open(path, 'r', encoding='utf-8') as f:
                                r = list(csv.DictReader(f))
                                if r:
                                    vmr = r[0]
                        except Exception:
                            vmr = None
                if vmr:
                    # fill mappings
                    map_keys = {
                        'Điểm A': ['DiemA','DiemA.Text','Diem A'],
                        'Điểm B': ['DiemB','DiemB.Text','Diem B'],
                        'Lượt cơ': ['Lco','Lco.Text','LuotCo.Text','LuotCo'],
                        'HR1A': ['HR1A','HR1A.Text'],'HR2A': ['HR2A','HR2A.Text'],
                        'HR1B': ['HR1B','HR1B.Text'],'HR2B': ['HR2B','HR2B.Text'],
                        'Tên VĐV A': ['TenA','TenA.Text'],'Tên VĐV B': ['TenB','TenB.Text'],
                        'Trận': ['Tran.Text', 'Tran', 'Trận']
                    }
                    for fld, keys in map_keys.items():
                        for k in keys:
                            if k in vmr and vmr.get(k) is not None:
                                try:
                                    entries[fld].delete(0,'end')
                                    entries[fld].insert(0, vmr.get(k))
                                except Exception:
                                    pass
                                break
                    compute_avgs()
                    return True
                return False

            def send_to_vmix():
                try:
                    import requests
                    vmix_url = self.match_rows[idx][5].get().strip() if hasattr(self, 'match_rows') and idx < len(self.match_rows) else ''
                    if not vmix_url:
                        messagebox.showerror('Lỗi', 'Không có địa chỉ vMix cho dòng này')
                        return False
                    mapping = {
                        'Tên VĐV A': 'TenA.Text','Tên VĐV B':'TenB.Text',
                        'Điểm A':'DiemA.Text','Điểm B':'DiemB.Text','Lượt cơ':'Lco.Text',
                        'HR1A':'HR1A.Text','HR2A':'HR2A.Text','HR1B':'HR1B.Text','HR2B':'HR2B.Text',
                        'AVGA':'AvgA.Text','AVGB':'AvgB.Text','Trận':'Tran.Text'
                    }
                    session = requests.Session()
                    for fld, sel in mapping.items():
                        val = entries[fld].get() if entries.get(fld) else ''
                        try:
                            session.get(f"{vmix_url}/API/", params={'Function':'SetText','Input':1,'SelectedName':sel,'Value':val}, timeout=2)
                        except Exception:
                            pass
                    messagebox.showinfo('Thành công', 'Đã gửi lên vMix')
                    return True
                except Exception as ex:
                    messagebox.showerror('Lỗi', f'Không gửi được: {ex}')
                    return False

            # Buttons (always visible at bottom)
            btn_frame = tk.Frame(popup, bg='#162033')
            btn_frame.pack(side='bottom', fill='x', padx=12, pady=(8,12))
            btn_inner = tk.Frame(btn_frame, bg='#162033')
            btn_inner.pack(anchor='center')
            tk.Button(
                btn_inner,
                text='Lấy từ vMix',
                command=prefill_from_vmix,
                bg='#2F6FED',
                fg='white',
                activebackground='#2459BC',
                activeforeground='white',
                relief='flat',
                bd=0,
                cursor='hand2',
                font=('Segoe UI', 13, 'bold')
            ).pack(side='left', padx=8, ipadx=12, ipady=7)
            # send and close helper binds Enter and button
            def send_and_close(event=None):
                ok = send_to_vmix()
                if not ok:
                    return
                try:
                    popup.destroy()
                except Exception:
                    pass

            tk.Button(
                btn_inner,
                text='Gửi',
                command=send_and_close,
                bg='#15A05A',
                fg='white',
                activebackground='#0F7A44',
                activeforeground='white',
                relief='flat',
                bd=0,
                cursor='hand2',
                font=('Segoe UI', 13, 'bold')
            ).pack(side='left', padx=8, ipadx=12, ipady=7)
            tk.Button(
                btn_inner,
                text='Hủy',
                command=popup.destroy,
                bg='#E45757',
                fg='white',
                activebackground='#B63F3F',
                activeforeground='white',
                relief='flat',
                bd=0,
                cursor='hand2',
                font=('Segoe UI', 13, 'bold')
            ).pack(side='left', padx=8, ipadx=12, ipady=7)

            # bind Enter to send
            try:
                popup.bind('<Return>', send_and_close)
                popup.bind('<KP_Enter>', send_and_close)
            except Exception:
                pass

            try:
                form_canvas.yview_moveto(0)
            except Exception:
                pass

            # auto-prefill
            try:
                popup.after(200, prefill_from_vmix)
            except Exception:
                prefill_from_vmix()
        except Exception as e:
            try:
                messagebox.showerror('Lỗi', f'Không mở được popup: {e}')
            except Exception:
                pass

    def _auto_restore_state_to_ui(self):
        """Gán lại giá trị vào UI từ self._restored_state nếu có."""
        s = getattr(self, '_restored_state', None)

        if not s:
            try:
                import pickle
                s = None
                for p in self._iter_state_candidates():
                    try:
                        if p and os.path.exists(p):
                            with open(p, 'rb') as f:
                                raw_state = pickle.load(f)
                                s = self._normalize_state_dict(raw_state)
                                if not s:
                                    continue
                                self._restored_state = s
                                break
                    except Exception:
                        continue
            except Exception:
                s = None

        # Không có state: giữ nguyên default hiện tại của UI.
        if not s:
            return

        try:
            self._restore_committed = False
        except Exception:
            pass

        try:
            import time
            self._autosave_suspended_until = time.time() + 5.5
        except Exception:
            pass

        try:
            if hasattr(self, 'tengiai_var') and s.get('tengiai') is not None:
                try:
                    self.tengiai_var.set(s.get('tengiai', ''))
                except Exception:
                    pass

            if hasattr(self, 'thoigian_var') and s.get('thoigian') is not None:
                try:
                    self.thoigian_var.set(s.get('thoigian', ''))
                except Exception:
                    pass

            if hasattr(self, 'diadiem_var') and s.get('diadiem') is not None:
                try:
                    self.diadiem_var.set(s.get('diadiem', ''))
                except Exception:
                    pass

            if hasattr(self, 'chuchay_var') and s.get('chuchay') is not None:
                try:
                    self.chuchay_var.set(s.get('chuchay', ''))
                except Exception:
                    pass

            if hasattr(self, 'chuchay_text') and s.get('chuchay') is not None:
                try:
                    self.chuchay_text.delete('1.0', 'end')
                    self.chuchay_text.insert('1.0', s.get('chuchay', ''))
                except Exception:
                    pass

            if hasattr(self, 'diemso_text') and s.get('diemso') is not None:
                try:
                    val = s.get('diemso', '')
                    self.diemso_text.delete('1.0', 'end')
                    self.diemso_text.insert('1.0', val)
                    try:
                        self.propagate_master_score(val)
                    except Exception:
                        pass
                except Exception:
                    pass

            if hasattr(self, 'url_var') and s.get('sheet_url') is not None:
                try:
                    self.url_var.set(s.get('sheet_url', ''))
                except Exception:
                    pass

            # Khôi phục các trường HBSF
            try:
                if s.get('hbsf_url'):
                    if hasattr(self, 'hbsf_url_var'):
                        self.hbsf_url_var.set(s['hbsf_url'])
                    else:
                        self._saved_hbsf_url = s['hbsf_url']
                if s.get('event_id') is not None:
                    if hasattr(self, 'event_id_var'):
                        self.event_id_var.set(str(s.get('event_id', '')))
                    else:
                        self._saved_event_id = str(s.get('event_id', ''))
                if s.get('round_type'):
                    if hasattr(self, 'round_type_var'):
                        self.round_type_var.set(s['round_type'])
                    else:
                        self._saved_round_type = s['round_type']
            except Exception:
                pass

            try:
                restored_creds = self._resolve_credentials_path(s.get('creds_path'))
                if restored_creds:
                    self.creds_path = restored_creds
                    if hasattr(self, 'creds_label'):
                        self.creds_label.config(text=os.path.basename(self.creds_path), fg='#00FF00')
                else:
                    self.creds_path = None
                    if hasattr(self, 'creds_label'):
                        self.creds_label.config(text='Chưa chọn credentials.json', fg='#FFB74D')
            except Exception:
                pass

            try:
                self._preview_footer_logo_path = str(s.get('preview_footer_logo') or '').strip()
            except Exception:
                self._preview_footer_logo_path = ''
            try:
                self._preview_meta_user_configured = bool(s.get('preview_user_configured', False))
            except Exception:
                self._preview_meta_user_configured = False

            try:
                ban_val = s.get('ban')
                if ban_val is None:
                    ban_val = len(s.get('table', []))
                if ban_val is not None and hasattr(self, 'ban_var') and self.ban_var is not None:
                    self.ban_var.set(ban_val)
            except Exception:
                pass

            try:
                self._restoring_state = True
            except Exception:
                pass

            try:
                self.populate_table()
            except Exception:
                pass

            table = s.get('table', [])
            if table:
                try:
                    self._apply_restored_table(table)
                except Exception:
                    pass

            try:
                self._restoring_state = False
            except Exception:
                pass

            if hasattr(self, 'diemso_text') and s.get('diemso') is not None:
                try:
                    val = s.get('diemso', '')
                    self.diemso_text.delete('1.0', 'end')
                    self.diemso_text.insert('1.0', val)
                    self.propagate_master_score(val)
                except Exception:
                    pass

            preview_data = s.get('preview', None)
            if preview_data:
                try:
                    self._last_preview_meta = []
                    for cell in preview_data:
                        try:
                            self._last_preview_meta.append({
                                'type': cell.get('type'),
                                'value': cell.get('value'),
                                'image_mode': cell.get('image_mode', 'fit'),
                                'logo_effect': cell.get('logo_effect', 'cut'),
                                'logo_interval': cell.get('logo_interval', 4.0),
                            })
                        except Exception:
                            self._last_preview_meta.append(None)
                except Exception:
                    pass
            else:
                try:
                    rows = len(getattr(self, 'match_rows', []) or [])
                    if rows > 0 and not getattr(self, '_last_preview_meta', None):
                        sp = []
                        for i in range(9):
                            if i < rows:
                                sp.append({'type': 'vmix', 'value': i, 'image_mode': 'fit'})
                            else:
                                sp.append(None)
                        try:
                            self._last_preview_meta = sp
                        except Exception:
                            pass
                        try:
                            self._auto_save_state()
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

        try:
            import time
            self._autosave_suspended_until = time.time() + 6.0
        except Exception:
            pass

        try:
            self._restore_committed = True
        except Exception:
            pass

        try:
            if s.get('_state_foreign_machine') and hasattr(self, 'status_var'):
                self.status_var.set('Đã nạp state từ máy khác: vui lòng kiểm tra lại URL vMix/credentials.')
        except Exception:
            pass

        self._auto_save_state()

    # --- Programmatic preview control helpers ---
    def preview_open(self):
        """Mở cửa sổ preview (gọi action của nút nếu cần). Trả về đối tượng preview hoặc None."""
        # If preview already exists and is valid, just deiconify
        p = getattr(self, '_preview_window', None)
        if p is not None:
            if p.winfo_exists():
                p.deiconify()
            print('[DEBUG] preview_open: preview_all_btn exists?', hasattr(self, 'preview_all_btn'))
            if hasattr(self, 'preview_all_btn'):
                self.preview_all_btn.invoke()
                print('[DEBUG] preview_open: invoked preview_all_btn')
            elif hasattr(self, 'open_preview_all'):
                self.open_preview_all()
                print('[DEBUG] preview_open: called open_preview_all fallback')
                return getattr(self, '_preview_window', None)

    def preview_set_cell(self, idx, kind, value, image_mode='fit', logo_effect='cut', logo_interval=4.0):
        """Cấu hình ô preview theo mã.
        kind: 'image' | 'vmix' | 'logo_playlist' | 'clear'
        value: đường dẫn ảnh, URL / row index, hoặc list đường dẫn logo
        image_mode: 'fit'|'center'|'cover' (chỉ với ảnh)
        Trả về True nếu đã set và kích hoạt render, False nếu lỗi.
        """
        try:
            p = getattr(self, '_preview_window', None)
            if p is None or not getattr(p, 'winfo_exists', lambda: False)():
                p = self.preview_open()
            if p is None:
                return False
            # ensure cell_meta exists
            try:
                meta = p.cell_meta
            except Exception:
                p.cell_meta = [{'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'} for _ in range(9)]
                meta = p.cell_meta
            if idx < 0 or idx >= len(meta):
                return False
            if kind == 'image':
                meta[idx] = {'type': 'image', 'value': value, 'image_ref': None, 'image_mode': image_mode}
            elif kind == 'logo_playlist':
                logos = []
                try:
                    if isinstance(value, (list, tuple)):
                        logos = [str(x).strip() for x in value if str(x).strip()]
                    elif isinstance(value, str) and value.strip():
                        logos = [value.strip()]
                except Exception:
                    logos = []
                try:
                    interval = float(logo_interval)
                except Exception:
                    interval = 4.0
                interval = max(0.5, min(120.0, interval))
                eff = str(logo_effect or 'cut').strip().lower()
                if eff not in ('cut', 'zoom', 'slide', 'cube', 'fly', 'mix'):
                    eff = 'cut'
                meta[idx] = {
                    'type': 'logo_playlist',
                    'value': logos,
                    'image_ref': None,
                    'image_mode': 'fit',
                    'logo_effect': eff,
                    'logo_interval': interval,
                }
            elif kind in ('clear', 'none', None):
                meta[idx] = {'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'}
            else:
                meta[idx] = {'type': 'vmix', 'value': value, 'image_ref': None}
            try:
                p._render_signatures[idx] = None
            except Exception:
                pass
            # trigger the <Configure> bound to the frame to force a redraw
            try:
                f = p.cells[idx]
                f.event_generate('<Configure>')
            except Exception:
                pass
            try:
                # persist last-known preview meta even if preview window closed later
                self._update_last_preview_meta(p)
            except Exception:
                pass
            return True
        except Exception:
            return False
    def open_preview_all(self):
        """Open a fullscreen 3x3 preview grid (standalone implementation).
        Each cell can show a vMix scoreboard (from a match row or URL) or an image.
        This method is self-contained so it can be invoked programmatically.
        """
        try:
            import requests
            import xml.etree.ElementTree as ET
            import threading
            import time
            import random
            from tkinter import filedialog, simpledialog, messagebox
            try:
                from PIL import Image, ImageTk
            except Exception:
                Image = ImageTk = None

            # Only allow one preview window
            if hasattr(self, '_preview_window') and self._preview_window is not None:
                try:
                    if self._preview_window.winfo_exists():
                        self._preview_window.deiconify()
                        self._preview_window.lift()
                        self._preview_window.focus_force()
                        return self._preview_window
                except Exception:
                    pass
            preview = tk.Toplevel(self)
            self._preview_window = preview
            preview._closing = False
            preview.title('Preview 3x3 Bảng điểm')

            def preview_alive():
                try:
                    return getattr(self, '_preview_window', None) is preview and preview.winfo_exists() and not getattr(preview, '_closing', False)
                except Exception:
                    return False

            def on_close():
                try:
                    preview._closing = True
                except Exception:
                    pass
                try:
                    self._preview_window = None
                except Exception:
                    pass
                try:
                    preview.destroy()
                except Exception:
                    pass
            preview.protocol('WM_DELETE_WINDOW', on_close)
            try:
                preview.attributes('-fullscreen', True)
            except Exception:
                try:
                    preview.state('zoomed')
                except Exception:
                    pass

            preview.cells = []
            preview._render_signatures = [None for _ in range(9)]
            preview._vmix_cache = {}
            preview._vmix_cache_lock = threading.Lock()
            preview._cache_ttl_sec = 1.8
            preview._error_retry_sec = 2.5
            preview.footer_logo_path = str(getattr(self, '_preview_footer_logo_path', '') or '')
            preview._footer_logo_cache = {'path': None, 'size': None, 'tkimg': None}
            preview._logo_playlist_states = {}
            preview._logo_image_cache = {}

            def _new_cache_entry():
                return {'data': None, 'error': None, 'updated_at': 0.0, 'loading': False, 'version': 0}

            def normalize_vmix_url(raw):
                try:
                    url = str(raw or '').strip()
                except Exception:
                    url = ''
                if not url:
                    return ''
                url = url.rstrip('/')
                if url.lower().endswith('/api'):
                    url = url[:-4].rstrip('/')
                return url

            def extract_input1_fields(xml_text):
                try:
                    root = ET.fromstring(xml_text)
                    input1 = root.find(".//input[@number='1']")
                    if input1 is None:
                        return None, 'Không tìm thấy Input 1 trong vMix'

                    text_nodes = list(input1.findall('text'))

                    def _norm_key(value):
                        try:
                            s = str(value or '').strip().lower()
                        except Exception:
                            return ''
                        for ch in (' ', '.', '_', '-'):
                            s = s.replace(ch, '')
                        return s

                    def get_field(name):
                        target = str(name or '').strip()
                        target_norm = _norm_key(target)
                        for text in text_nodes:
                            raw_name = str(text.attrib.get('name') or '').strip()
                            if raw_name == target:
                                return text.text or ''
                            if raw_name.lower() == target.lower():
                                return text.text or ''
                            if _norm_key(raw_name) == target_norm:
                                return text.text or ''
                        return ''

                    def get_first(*names):
                        for name in names:
                            value = get_field(name)
                            if value:
                                return value
                        return ''

                    noidung_1 = get_first(
                        'Noidung.Text', 'NoiDung.Text', 'Noi dung.Text',
                        'NOIDUNG.Text', 'NOI DUNG.Text',
                        'Noidung.txt', 'NoiDung.txt', 'Noi dung.txt',
                        'Noidung', 'NoiDung', 'Noi dung'
                    )
                    noidung_2 = get_first(
                        'Noidung2.Text', 'NoiDung2.Text', 'Noi dung2.Text',
                        'Noidung1.Text', 'NoiDung1.Text', 'Noi dung1.Text',
                        'Noidung2', 'NoiDung2', 'Noi dung2'
                    )
                    noidung_combined = noidung_1
                    if noidung_2:
                        noidung_combined = f"{noidung_combined}\n{noidung_2}".strip() if noidung_combined else noidung_2

                    data = {
                        'ten_a': get_first('TenA.Text', 'TenA'),
                        'ten_b': get_first('TenB.Text', 'TenB'),
                        'tran': get_first('Tran.Text', 'Tran', 'Trận.Text', 'Trận'),
                        'noidung': noidung_combined,
                        'diem_a': get_first('DiemA.Text', 'DiemA'),
                        'diem_b': get_first('DiemB.Text', 'DiemB'),
                        'avg_a': get_first('AvgA.Text', 'AVGA', 'AvgA'),
                        'avg_b': get_first('AvgB.Text', 'AVGB', 'AvgB'),
                        'lco': get_first('Lco.Text', 'Lco', 'LuotCo.Text', 'LuotCo'),
                        'adi': get_first('Adi.Text', 'ADI.Text', 'Adi', 'ADI'),
                        'bdi': get_first('Bdi.Text', 'BDI.Text', 'Bdi', 'BDI'),
                        'ta': get_first('TA.Text', 'Ta.Text', 'ta.text', 'TA', 'Ta', 'ta'),
                        'tb': get_first('TB.Text', 'Tb.Text', 'tb.text', 'TB', 'Tb', 'tb'),
                        't1_source': get_first('T1.Source', 'T1.Source.Text', 't1.source', 't1.source.text', 'T1Source', 'T1SOURCE', 'T1'),
                        't2_source': get_first('T2.Source', 'T2.Source.Text', 't2.source', 't2.source.text', 'T2Source', 'T2SOURCE', 'T2'),
                        't3_source': get_first('T3.Source', 'T3.Source.Text', 't3.source', 't3.source.text', 'T3Source', 'T3SOURCE', 'T3'),
                        't4_source': get_first('T4.Source', 'T4.Source.Text', 't4.source', 't4.source.text', 'T4Source', 'T4SOURCE', 'T4'),
                        'hr1_a': get_first('HR1A.Text', 'HR1A'),
                        'hr2_a': get_first('HR2A.Text', 'HR2A'),
                        'hr1_b': get_first('HR1B.Text', 'HR1B'),
                        'hr2_b': get_first('HR2B.Text', 'HR2B'),
                        'timer': get_first('Time.Text', 'Timer.Text', 'Clock.Text', 'Time', 'Timer', 'Clock'),
                        'ban': get_first('Ban.Text', 'Ban', 'Bàn.Text', 'Bàn'),
                    }
                    return data, None
                except Exception as ex:
                    return None, str(ex)

            def resolve_vmix_url(cell_meta):
                try:
                    raw_value = cell_meta.get('value')
                    if isinstance(raw_value, int):
                        rowidx = raw_value
                        if rowidx < len(getattr(self, 'match_rows', [])):
                            row = self.match_rows[rowidx]
                            raw_value = row[5].get().strip() if hasattr(row[5], 'get') else ''
                        else:
                            raw_value = ''
                    return normalize_vmix_url(raw_value)
                except Exception:
                    return ''

            def resolve_row_index(cell_meta):
                try:
                    raw_value = cell_meta.get('value')
                    if isinstance(raw_value, int):
                        rowidx = int(raw_value)
                        if 0 <= rowidx < len(getattr(self, 'match_rows', [])):
                            return rowidx
                        return None

                    vmix_url = normalize_vmix_url(raw_value)
                    if not vmix_url:
                        return None
                    for ridx, row in enumerate(getattr(self, 'match_rows', []) or []):
                        try:
                            row_url = normalize_vmix_url(row[5].get().strip() if len(row) > 5 and hasattr(row[5], 'get') else '')
                        except Exception:
                            row_url = ''
                        if row_url and row_url == vmix_url:
                            return ridx
                except Exception:
                    return None
                return None

            def resolve_table_label(cell_meta, vmix_data=None):
                try:
                    if isinstance(vmix_data, dict):
                        from_vmix = str(vmix_data.get('ban') or '').strip()
                        if from_vmix:
                            return from_vmix
                except Exception:
                    pass

                try:
                    ridx = resolve_row_index(cell_meta)
                    if ridx is not None and ridx < len(getattr(self, 'match_rows', [])):
                        row = self.match_rows[ridx]
                        if len(row) > 1 and hasattr(row[1], 'get'):
                            table_val = str(row[1].get() or '').strip()
                            if table_val:
                                return table_val
                except Exception:
                    pass
                return ''

            def build_vmix_fingerprint(vmix_data):
                if not isinstance(vmix_data, dict):
                    return None
                keys = (
                    'ten_a', 'ten_b', 'tran', 'noidung',
                    'diem_a', 'diem_b',
                    'avg_a', 'avg_b',
                    'hr1_a', 'hr2_a', 'hr1_b', 'hr2_b',
                    'lco', 'adi', 'bdi',
                    'ta', 'tb', 't1_source', 't2_source', 't3_source', 't4_source',
                    'timer', 'ban'
                )
                out = []
                try:
                    for k in keys:
                        out.append((k, str(vmix_data.get(k, '') or '').strip()))
                except Exception:
                    return None
                return tuple(out)

            def get_cache_snapshot(vmix_url):
                if not vmix_url:
                    return None
                try:
                    with preview._vmix_cache_lock:
                        cur = preview._vmix_cache.get(vmix_url)
                        if cur is None:
                            return None
                        return dict(cur)
                except Exception:
                    return None

            def schedule_vmix_fetch(vmix_url, force=False):
                vmix_url = normalize_vmix_url(vmix_url)
                if not vmix_url or not preview_alive():
                    return
                now = time.monotonic()
                try:
                    with preview._vmix_cache_lock:
                        entry = preview._vmix_cache.get(vmix_url)
                        if entry is None:
                            entry = _new_cache_entry()
                            preview._vmix_cache[vmix_url] = entry

                        age = now - float(entry.get('updated_at', 0.0) or 0.0)
                        has_data = bool(entry.get('data'))
                        has_error = bool(entry.get('error'))
                        ttl = preview._cache_ttl_sec if has_data else (preview._error_retry_sec if has_error else 0.0)

                        if entry.get('loading'):
                            return
                        if not force and ttl > 0 and age < ttl:
                            return
                        entry['loading'] = True
                except Exception:
                    return

                def _worker(url=vmix_url):
                    parsed = None
                    error_text = None
                    try:
                        resp = requests.get(f'{url}/API/', timeout=(0.8, 1.5))
                        resp.raise_for_status()
                        parsed, parse_error = extract_input1_fields(resp.text)
                        if parse_error:
                            error_text = parse_error
                    except Exception as ex:
                        error_text = str(ex)

                    def _apply_result():
                        if not preview_alive():
                            return
                        try:
                            with preview._vmix_cache_lock:
                                cur = preview._vmix_cache.get(url)
                                if cur is None:
                                    cur = _new_cache_entry()
                                    preview._vmix_cache[url] = cur
                                cur['loading'] = False
                                cur['updated_at'] = time.monotonic()
                                cur['version'] = int(cur.get('version', 0) or 0) + 1
                                if parsed is not None:
                                    cur['data'] = parsed
                                    cur['error'] = None
                                else:
                                    cur['error'] = error_text or 'Không đọc được dữ liệu vMix'
                        except Exception:
                            return

                        for ii, meta in enumerate(getattr(preview, 'cell_meta', []) or []):
                            try:
                                if meta and meta.get('type') == 'vmix' and resolve_vmix_url(meta) == url:
                                    preview._render_signatures[ii] = None
                                    render_cell_local(ii)
                            except Exception:
                                pass

                    try:
                        self.after(0, _apply_result)
                    except Exception:
                        pass

                try:
                    threading.Thread(target=_worker, daemon=True).start()
                except Exception:
                    try:
                        with preview._vmix_cache_lock:
                            cur = preview._vmix_cache.get(vmix_url)
                            if cur is not None:
                                cur['loading'] = False
                    except Exception:
                        pass

            # Auto-map vMix/table rows to preview grid cells if match_rows is available
            num_tables = len(getattr(self, 'match_rows', []))
            # Default: all empty
            preview.cell_meta = [{'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'} for _ in range(9)]
            # Always map tables in natural order and keep center cell (index 4, ô số 5) empty.
            default_cell_order = [0, 1, 2, 3, 5, 6, 7, 8]
            for table_idx, cell_idx in enumerate(default_cell_order):
                if table_idx >= num_tables:
                    break
                preview.cell_meta[cell_idx] = {'type': 'vmix', 'value': table_idx, 'image_ref': None, 'image_mode': 'fit'}

            # Overlay only non-vMix manual cells from last snapshot so auto table allocation is preserved.
            try:
                saved_preview = getattr(self, '_last_preview_meta', None)
                if isinstance(saved_preview, list) and len(saved_preview) >= 1:
                    for i in range(9):
                        if i >= len(saved_preview):
                            break
                        cell = saved_preview[i]
                        if not isinstance(cell, dict):
                            continue
                        ctype = str(cell.get('type') or '').strip().lower()
                        if ctype == 'image':
                            preview.cell_meta[i] = {
                                'type': 'image',
                                'value': cell.get('value'),
                                'image_ref': None,
                                'image_mode': cell.get('image_mode', 'fit'),
                            }
                        elif ctype == 'logo_playlist':
                            raw_logo_value = cell.get('value')
                            logos = []
                            try:
                                if isinstance(raw_logo_value, (list, tuple)):
                                    logos = [str(x).strip() for x in raw_logo_value if str(x).strip()]
                                elif isinstance(raw_logo_value, str) and raw_logo_value.strip():
                                    txt = raw_logo_value.strip()
                                    if '|' in txt:
                                        logos = [s.strip() for s in txt.split('|') if s.strip()]
                                    else:
                                        logos = [txt]
                            except Exception:
                                logos = []
                            if logos:
                                preview.cell_meta[i] = {
                                    'type': 'logo_playlist',
                                    'value': logos,
                                    'image_ref': None,
                                    'image_mode': 'fit',
                                    'logo_effect': str(cell.get('logo_effect') or 'cut'),
                                    'logo_interval': cell.get('logo_interval', 4.0),
                                }
                        elif ctype == 'vmix' and (num_tables <= 0 or getattr(self, '_preview_meta_user_configured', False)):
                            # If table rows are not available yet, keep saved vMix mapping as fallback.
                            # Also respect user-configured cell assignments from 'Cấu hình Preview' dialog.
                            preview.cell_meta[i] = {
                                'type': 'vmix',
                                'value': cell.get('value'),
                                'image_ref': None,
                                'image_mode': 'fit',
                                'ban_name': cell.get('ban_name'),
                            }
            except Exception:
                pass

            def cell_size(frame):
                try:
                    fw = max(10, frame.winfo_width())
                    fh = max(10, frame.winfo_height())
                except Exception:
                    fw = preview.winfo_screenwidth() // 3
                    fh = preview.winfo_screenheight() // 3
                return fw, fh

            def get_footer_logo_tkimg(max_w, max_h):
                try:
                    path = str(getattr(preview, 'footer_logo_path', '') or '').strip()
                    if not path or not Image or not ImageTk or not os.path.exists(path):
                        return None
                    ww = max(8, int(max_w))
                    hh = max(8, int(max_h))
                    cache = getattr(preview, '_footer_logo_cache', None)
                    if isinstance(cache, dict):
                        if cache.get('path') == path and cache.get('size') == (ww, hh) and cache.get('tkimg') is not None:
                            return cache.get('tkimg')
                    with Image.open(path) as base_img:
                        img = base_img.copy()
                    iw, ih = img.size
                    if iw <= 0 or ih <= 0:
                        return None
                    try:
                        resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                    except Exception:
                        resample = Image.BICUBIC
                    scale = min(float(ww) / float(iw), float(hh) / float(ih))
                    nw = max(1, int(iw * scale))
                    nh = max(1, int(ih * scale))
                    img2 = img.resize((nw, nh), resample)
                    tkimg = ImageTk.PhotoImage(img2)
                    preview._footer_logo_cache = {'path': path, 'size': (ww, hh), 'tkimg': tkimg}
                    return tkimg
                except Exception:
                    return None

            def normalize_logo_paths(raw_value):
                paths = []
                try:
                    if isinstance(raw_value, (list, tuple)):
                        src = list(raw_value)
                    elif isinstance(raw_value, str):
                        txt = raw_value.strip()
                        if '|' in txt:
                            src = txt.split('|')
                        else:
                            src = [txt] if txt else []
                    else:
                        src = []
                    for item in src:
                        p = str(item or '').strip()
                        if p:
                            paths.append(p)
                except Exception:
                    paths = []
                return paths

            def normalize_logo_effect(raw_effect):
                eff = str(raw_effect or 'cut').strip().lower()
                if eff not in ('cut', 'zoom', 'slide', 'cube', 'fly', 'mix'):
                    eff = 'cut'
                return eff

            def normalize_logo_interval(raw_interval, default_value=4.0):
                try:
                    iv = float(raw_interval)
                except Exception:
                    iv = float(default_value)
                return max(0.5, min(120.0, iv))

            def get_playlist_base_image(path, max_w, max_h):
                try:
                    if not (Image and ImageTk):
                        return None
                    p = str(path or '').strip()
                    if not p or not os.path.exists(p):
                        return None
                    ww = max(8, int(max_w))
                    hh = max(8, int(max_h))
                    key = (p, ww, hh)
                    cache = getattr(preview, '_logo_image_cache', None)
                    if not isinstance(cache, dict):
                        cache = {}
                        preview._logo_image_cache = cache
                    img = cache.get(key)
                    if img is not None:
                        return img
                    with Image.open(p) as src:
                        base = src.convert('RGBA')
                    iw, ih = base.size
                    if iw <= 0 or ih <= 0:
                        return None
                    try:
                        resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                    except Exception:
                        resample = Image.BICUBIC
                    scale = min(float(ww) / float(iw), float(hh) / float(ih))
                    nw = max(1, int(iw * scale))
                    nh = max(1, int(ih * scale))
                    out = base.resize((nw, nh), resample)
                    cache[key] = out
                    if len(cache) > 220:
                        for old_key in list(cache.keys())[:40]:
                            cache.pop(old_key, None)
                    return out
                except Exception:
                    return None

            def transform_image_for_effect(img, sx=1.0, sy=1.0):
                try:
                    if img is None:
                        return None
                    sx = max(0.02, float(sx))
                    sy = max(0.02, float(sy))
                    nw = max(1, int(img.width * sx))
                    nh = max(1, int(img.height * sy))
                    if nw == img.width and nh == img.height:
                        return img
                    try:
                        resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                    except Exception:
                        resample = Image.BICUBIC
                    return img.resize((nw, nh), resample)
                except Exception:
                    return img

            def apply_image_alpha(img, alpha=1.0):
                try:
                    if img is None:
                        return None
                    a = max(0.0, min(1.0, float(alpha)))
                    if a >= 0.999:
                        return img
                    if a <= 0.001:
                        return None
                    out = img.copy()
                    if out.mode != 'RGBA':
                        out = out.convert('RGBA')
                    try:
                        channel = out.getchannel('A')
                        channel = channel.point(lambda px: int(px * a))
                        out.putalpha(channel)
                    except Exception:
                        pass
                    return out
                except Exception:
                    return img

            def draw_pil_on_canvas(canvas, refs, img, cx, cy, alpha=1.0):
                try:
                    if not (Image and ImageTk):
                        return
                    final_img = apply_image_alpha(img, alpha)
                    if final_img is None:
                        return
                    tkimg = ImageTk.PhotoImage(final_img)
                    refs.append(tkimg)
                    canvas.create_image(int(cx), int(cy), image=tkimg)
                except Exception:
                    pass

            def render_logo_playlist_cell(idx, frame, cell, w, h):
                logos = normalize_logo_paths(cell.get('value') if isinstance(cell, dict) else None)
                effect = normalize_logo_effect(cell.get('logo_effect') if isinstance(cell, dict) else 'cut')
                interval_sec = normalize_logo_interval(cell.get('logo_interval') if isinstance(cell, dict) else 4.0, 4.0)
                transition_sec = min(1.4, max(0.25, interval_sec * 0.35))

                cfg_key = (tuple(logos), effect, round(interval_sec, 2), round(transition_sec, 2), int(w), int(h))
                state_map = getattr(preview, '_logo_playlist_states', None)
                if not isinstance(state_map, dict):
                    state_map = {}
                    preview._logo_playlist_states = state_map
                st = state_map.get(idx)

                # Rebuild canvas only when config/size changed.
                if not isinstance(st, dict) or st.get('cfg_key') != cfg_key:
                    for child in list(frame.winfo_children()):
                        try:
                            child.destroy()
                        except Exception:
                            pass
                    canvas = tk.Canvas(frame, bg='#020202', bd=0, highlightthickness=0)
                    canvas.place(x=0, y=0, width=w, height=h)
                    now_init = time.monotonic()
                    st = {
                        'cfg_key': cfg_key,
                        'canvas': canvas,
                        'current': 0,
                        'next': 1 if len(logos) > 1 else 0,
                        'phase': 'hold',
                        'phase_started': now_init,
                        'last_switch': now_init,
                        'refs': [],
                    }
                    state_map[idx] = st
                else:
                    canvas = st.get('canvas')
                    try:
                        canvas.place(x=0, y=0, width=w, height=h)
                        canvas.configure(width=w, height=h)
                    except Exception:
                        pass

                # Nothing to show -> placeholder.
                if not logos:
                    try:
                        canvas.delete('all')
                        canvas.create_text(
                            int(w / 2),
                            int(h / 2),
                            text='CHƯA CÓ LOGO',
                            fill='#7A6740',
                            font=('Segoe UI', max(10, min(26, int(min(w, h) * 0.08))), 'bold')
                        )
                        st['refs'] = []
                    except Exception:
                        pass
                    return

                now = time.monotonic()
                total = len(logos)
                try:
                    st['current'] = int(st.get('current', 0)) % total
                except Exception:
                    st['current'] = 0

                progress = 0.0
                if total > 1:
                    if st.get('phase') == 'hold' and (now - float(st.get('last_switch', now))) >= interval_sec:
                        st['phase'] = 'transition'
                        st['phase_started'] = now
                        st['next'] = (int(st.get('current', 0)) + 1) % total
                        if effect == 'mix':
                            try:
                                st['active_effect'] = random.choice(('cut', 'zoom', 'slide', 'cube', 'fly'))
                            except Exception:
                                st['active_effect'] = 'cut'
                        else:
                            st['active_effect'] = effect
                    if st.get('phase') == 'transition':
                        progress = (now - float(st.get('phase_started', now))) / max(0.05, transition_sec)
                        if progress >= 1.0:
                            st['current'] = int(st.get('next', st.get('current', 0))) % total
                            st['phase'] = 'hold'
                            st['last_switch'] = now
                            progress = 1.0

                cur_idx = int(st.get('current', 0)) % total
                next_idx = int(st.get('next', cur_idx)) % total
                cur_img = get_playlist_base_image(logos[cur_idx], int(w * 0.92), int(h * 0.90))
                nxt_img = get_playlist_base_image(logos[next_idx], int(w * 0.92), int(h * 0.90)) if total > 1 else cur_img

                refs = []
                try:
                    canvas.delete('all')
                except Exception:
                    pass

                cx = float(w) / 2.0
                cy = float(h) / 2.0

                if cur_img is None:
                    try:
                        canvas.create_text(
                            int(cx),
                            int(cy),
                            text='LOGO KHÔNG HỢP LỆ',
                            fill='#B85757',
                            font=('Segoe UI', max(9, min(24, int(min(w, h) * 0.07))), 'bold')
                        )
                    except Exception:
                        pass
                    st['refs'] = refs
                    return

                p = max(0.0, min(1.0, float(progress)))
                in_transition = total > 1 and st.get('phase') == 'transition'
                active_effect = str(st.get('active_effect') or ('cut' if effect == 'mix' else effect)).strip().lower()

                if not in_transition:
                    draw_pil_on_canvas(canvas, refs, cur_img, cx, cy, alpha=1.0)
                else:
                    if active_effect == 'cut':
                        draw_pil_on_canvas(canvas, refs, nxt_img if (p >= 0.5 and nxt_img is not None) else cur_img, cx, cy, alpha=1.0)
                    elif active_effect == 'zoom':
                        draw_pil_on_canvas(canvas, refs, cur_img, cx, cy, alpha=max(0.35, 1.0 - (0.65 * p)))
                        if nxt_img is not None:
                            z = 0.62 + (0.38 * p)
                            draw_pil_on_canvas(canvas, refs, transform_image_for_effect(nxt_img, z, z), cx, cy, alpha=0.35 + (0.65 * p))
                    elif active_effect == 'slide':
                        draw_pil_on_canvas(canvas, refs, cur_img, cx - (float(w) * p), cy, alpha=1.0)
                        if nxt_img is not None:
                            draw_pil_on_canvas(canvas, refs, nxt_img, cx + (float(w) * (1.0 - p)), cy, alpha=1.0)
                    elif active_effect == 'fly':
                        draw_pil_on_canvas(canvas, refs, cur_img, cx, cy, alpha=max(0.35, 1.0 - (0.65 * p)))
                        if nxt_img is not None:
                            start_x = float(w) * 1.18
                            start_y = float(h) * 1.15
                            nx = start_x + (cx - start_x) * p
                            ny = start_y + (cy - start_y) * p
                            z = 0.55 + (0.45 * p)
                            draw_pil_on_canvas(canvas, refs, transform_image_for_effect(nxt_img, z, z), nx, ny, alpha=0.35 + (0.65 * p))
                    elif active_effect == 'cube':
                        old_sx = max(0.08, 1.0 - p)
                        new_sx = max(0.08, p)
                        draw_pil_on_canvas(canvas, refs, transform_image_for_effect(cur_img, old_sx, 1.0), cx - (float(w) * 0.18 * p), cy, alpha=max(0.35, 1.0 - (0.65 * p)))
                        if nxt_img is not None:
                            draw_pil_on_canvas(canvas, refs, transform_image_for_effect(nxt_img, new_sx, 1.0), cx + (float(w) * 0.18 * (1.0 - p)), cy, alpha=0.35 + (0.65 * p))
                    else:
                        draw_pil_on_canvas(canvas, refs, nxt_img if nxt_img is not None else cur_img, cx, cy, alpha=1.0)

                st['refs'] = refs

            def render_vmix_cell_like_sample(frame, data, table_label, w, h):
                try:
                    from tkinter import font as tkfont

                    base = tk.Frame(frame, bg='#000000')
                    base.place(x=0, y=0, relwidth=1, relheight=1)

                    def _fs(scale, min_v, max_v):
                        try:
                            # Normalize to a 16:9 reference box so 16:10 screens
                            # do not inflate fonts compared to 16:9 screens.
                            ref_h = min(float(h), float(w) * (9.0 / 16.0))
                            ref_w = min(float(w), float(h) * (16.0 / 9.0))
                            base_dim = max(1.0, min(ref_w, ref_h))
                            size = int(base_dim * scale)
                        except Exception:
                            size = min_v
                        return max(min_v, min(max_v, size))

                    def _fit_font_size(text, family, weight, start_size, max_width_px, min_size=7):
                        try:
                            txt = str(text or '')
                            sz = int(start_size)
                            while sz > min_size:
                                f = tkfont.Font(family=family, size=sz, weight=weight)
                                if f.measure(txt) <= max_width_px:
                                    return sz
                                sz -= 1
                            return min_size
                        except Exception:
                            return max(min_size, int(start_size))

                    def _fit_multiline_font_size(lines, family, weight, start_size, max_width_px, max_height_px, min_size=7):
                        try:
                            clean = [str(x) for x in (lines or [])]
                            if not clean:
                                return min_size
                            sz = int(start_size)
                            while sz > min_size:
                                f = tkfont.Font(family=family, size=sz, weight=weight)
                                line_h = f.metrics('linespace')
                                total_h = line_h * len(clean)
                                max_w = 0
                                for ln in clean:
                                    mw = f.measure(ln)
                                    if mw > max_w:
                                        max_w = mw
                                if max_w <= max_width_px and total_h <= max_height_px:
                                    return sz
                                sz -= 1
                            return min_size
                        except Exception:
                            return max(min_size, int(start_size))

                    name_a = str(data.get('ten_a') or '').strip() or 'VĐV A'
                    name_b = str(data.get('ten_b') or '').strip() or 'VĐV B'
                    tran_raw = str(data.get('tran') or '').strip() or '--'
                    diem_a = str(data.get('diem_a') or '').strip() or '0'
                    diem_b = str(data.get('diem_b') or '').strip() or '0'
                    avg_a = str(data.get('avg_a') or '').strip() or '-'
                    avg_b = str(data.get('avg_b') or '').strip() or '-'
                    hr1_a = str(data.get('hr1_a') or '').strip() or '-'
                    hr2_a = str(data.get('hr2_a') or '').strip() or '-'
                    hr1_b = str(data.get('hr1_b') or '').strip() or '-'
                    hr2_b = str(data.get('hr2_b') or '').strip() or '-'
                    lco = str(data.get('lco') or '').strip() or '-'
                    noidung_raw = str(data.get('noidung') or '').strip()
                    adi = str(data.get('adi') or '').strip()
                    bdi = str(data.get('bdi') or '').strip()
                    ta_raw = str(data.get('ta') or '').strip()
                    tb_raw = str(data.get('tb') or '').strip()
                    t1_source = str(data.get('t1_source') or '').strip()
                    t2_source = str(data.get('t2_source') or '').strip()
                    t3_source = str(data.get('t3_source') or '').strip()
                    t4_source = str(data.get('t4_source') or '').strip()
                    timer = str(data.get('timer') or '').strip()

                    if noidung_raw:
                        noidung_text = noidung_raw.replace('\r\n', '\n').replace('\r', '\n')
                    else:
                        noidung_text = ''

                    tran_lower = tran_raw.lower()
                    if tran_lower.startswith('trận ') or tran_lower.startswith('tran '):
                        tran_suffix = tran_raw.split(' ', 1)[1] if ' ' in tran_raw else ''
                        tran_text = f'Trận {tran_suffix}'.strip()
                    elif tran_raw == '--':
                        tran_text = 'Trận --'
                    else:
                        tran_text = f'Trận {tran_raw}'

                    table_text = str(table_label or '').strip()
                    if table_text:
                        up = table_text.upper()
                        if not up.startswith('BÀN') and not up.startswith('BAN'):
                            table_text = f'BÀN {table_text}'
                        else:
                            table_text = up
                    else:
                        table_text = 'BÀN --'

                    header = tk.Frame(base, bg='#0A0A0A')
                    header.place(relx=0.0, rely=0.0, relwidth=1.0, relheight=0.11)
                    name_fs_base = _fs(0.052, 8, 20)
                    left_w = int(w * 0.45)
                    right_w = int(w * 0.45)
                    name_a_fs = _fit_font_size(name_a.upper(), 'Segoe UI', 'bold', name_fs_base, left_w, min_size=7)
                    name_b_fs = _fit_font_size(name_b.upper(), 'Segoe UI', 'bold', name_fs_base, right_w, min_size=7)
                    tk.Label(header, text=name_a.upper(), fg='#F0F0F0', bg='#0A0A0A', anchor='w', font=('Segoe UI', name_a_fs, 'bold')).place(relx=0.02, rely=0.08, relwidth=0.46, relheight=0.84)
                    tk.Label(header, text=name_b.upper(), fg='#F2B705', bg='#0A0A0A', anchor='e', font=('Segoe UI', name_b_fs, 'bold')).place(relx=0.52, rely=0.08, relwidth=0.46, relheight=0.84)

                    score_area = tk.Frame(base, bg='#000000')
                    score_area.place(relx=0.0, rely=0.11, relwidth=1.0, relheight=0.38)
                    left_box = tk.Frame(score_area, bg='#050505', highlightthickness=1, highlightbackground='#1B1B1B')
                    right_box = tk.Frame(score_area, bg='#050505', highlightthickness=1, highlightbackground='#1B1B1B')
                    left_box.place(relx=0.07, rely=0.04, relwidth=0.34, relheight=0.92)
                    right_box.place(relx=0.59, rely=0.04, relwidth=0.34, relheight=0.92)

                    center_col = tk.Frame(score_area, bg='#000000')
                    center_col.place(relx=0.41, rely=0.04, relwidth=0.18, relheight=0.92)

                    score_fs_base = _fs(0.185, 18, 88)
                    score_fs_a = _fit_font_size(diem_a, 'Arial', 'bold', score_fs_base, int(w * 0.33), min_size=16)
                    score_fs_b = _fit_font_size(diem_b, 'Arial', 'bold', score_fs_base, int(w * 0.33), min_size=16)
                    tk.Label(left_box, text=diem_a, fg='#E2E2E2', bg='#050505', font=('Arial', score_fs_a, 'bold')).place(relx=0.5, rely=0.5, anchor='center')
                    tk.Label(right_box, text=diem_b, fg='#F4B000', bg='#050505', font=('Arial', score_fs_b, 'bold')).place(relx=0.5, rely=0.5, anchor='center')

                    def _source_enabled(value):
                        s = str(value or '').strip().lower()
                        # NOTE: "0" can be a valid vMix source index, so it must not be treated as disabled.
                        return s not in ('', 'false', 'off', 'none', 'null')

                    def _turn_count(turn_value):
                        sval = str(turn_value or '').strip()
                        if not sval:
                            return None
                        try:
                            turns = int(float(sval.replace(',', '.')))
                        except Exception:
                            return None
                        if turns <= 0:
                            return 2
                        if turns == 1:
                            return 1
                        return 0

                    def _clock_count(turn_value, score_value, src_values):
                        source_values = list(src_values or [])
                        source_specified = any(str(v or '').strip() != '' for v in source_values)
                        source_count = sum(1 for v in source_values if _source_enabled(v))
                        has_score = str(score_value or '').strip() not in ('', '-')

                        turn_count = _turn_count(turn_value)
                        # Highest priority: explicit active source fields from vMix.
                        if source_specified and source_count > 0:
                            return source_count

                        # Next: TA/TB turns if they indicate an active clock count (>0).
                        if turn_count is not None and turn_count > 0:
                            return turn_count

                        # Safety fallback: if this scoreboard side still has score data,
                        # keep clocks visible to avoid "missing clocks" on some tables.
                        return 2 if has_score else 0

                    left_clock_count = _clock_count(ta_raw, diem_a, [t1_source, t2_source])
                    right_clock_count = _clock_count(tb_raw, diem_b, [t3_source, t4_source])

                    def draw_side_clocks(parent, side='left', count=0):
                        try:
                            total = max(0, min(2, int(count)))
                            if total <= 0:
                                return
                            size = max(16, min(40, int(min(w, h) * 0.085)))
                            relx = 0.058 if side == 'left' else 0.942
                            positions = [0.38] if total == 1 else [0.24, 0.52]
                            for rely in positions:
                                c = tk.Canvas(parent, width=size, height=size, bg='#000000', bd=0, highlightthickness=0)
                                c.place(relx=relx, rely=rely, anchor='center')
                                c.create_oval(1, 1, size - 1, size - 1, fill='#E2B11E', outline='#D39A00', width=1)
                                cx = size / 2
                                cy = size / 2
                                c.create_line(cx, cy, cx, cy - (size * 0.28), fill='#1A1A1A', width=max(1, int(size * 0.07)))
                                c.create_line(cx, cy, cx + (size * 0.18), cy + (size * 0.08), fill='#1A1A1A', width=max(1, int(size * 0.07)))
                        except Exception:
                            pass

                    draw_side_clocks(score_area, side='left', count=left_clock_count)
                    draw_side_clocks(score_area, side='right', count=right_clock_count)

                    if noidung_text:
                        noidung_lines = [ln.strip().upper() for ln in noidung_text.split('\n') if ln.strip()]
                        if not noidung_lines:
                            noidung_lines = [noidung_text.upper()]
                        noidung_fs = _fit_multiline_font_size(
                            noidung_lines,
                            'Segoe UI',
                            'bold',
                            _fs(0.094, 14, 40),
                            int(w * 0.205),
                            int(h * 0.28),
                            min_size=9
                        )
                        tk.Label(
                            center_col,
                            text='\n'.join(noidung_lines),
                            fg='#D9D9D9',
                            bg='#000000',
                            justify='center',
                            wraplength=max(52, int(w * 0.19)),
                            font=('Segoe UI', noidung_fs, 'bold')
                        ).place(relx=0.5, rely=0.48, anchor='center')

                    if timer:
                        timer_fs = _fs(0.088, 12, 46)
                        tk.Label(score_area, text=timer, fg='#F4B000', bg='#000000', font=('Consolas', timer_fs, 'bold')).place(relx=0.5, rely=0.87, anchor='center')

                    tk.Frame(base, bg='#7A3A3A').place(relx=0.0, rely=0.49, relwidth=1.0, relheight=0.005)

                    stats_area = tk.Frame(base, bg='#000000')
                    stats_area.place(relx=0.0, rely=0.50, relwidth=1.0, relheight=0.34)
                    left_stats = f'AVG {avg_a}\nHR1   {hr1_a}\nHR2   {hr2_a}'
                    right_stats = f'AVG {avg_b}\nHR1   {hr1_b}\nHR2   {hr2_b}'
                    left_lines = [f'AVG {avg_a}', f'HR1   {hr1_a}', f'HR2   {hr2_a}']
                    right_lines = [f'AVG {avg_b}', f'HR1   {hr1_b}', f'HR2   {hr2_b}']
                    stats_px_h = max(30, int(h * 0.34 * 0.84))
                    stats_lane_h = max(1, int(h * 0.34))
                    left_stats_fs = _fit_multiline_font_size(left_lines, 'Segoe UI', 'bold', _fs(0.056, 11, 26), int(w * 0.22), stats_px_h, min_size=8)
                    right_stats_fs = _fit_multiline_font_size(right_lines, 'Segoe UI', 'bold', _fs(0.056, 11, 26), int(w * 0.22), stats_px_h, min_size=8)

                    left_stats_box = tk.Frame(stats_area, bg='#000000')
                    right_stats_box = tk.Frame(stats_area, bg='#000000')
                    left_stats_box.place(relx=0.02, rely=0.04, relwidth=0.24, relheight=0.92)
                    right_stats_box.place(relx=0.74, rely=0.04, relwidth=0.24, relheight=0.92)
                    tk.Label(left_stats_box, text=left_stats, fg='#DCDCDC', bg='#000000', justify='left', anchor='w', font=('Segoe UI', left_stats_fs, 'bold')).pack(fill='both', expand=True)
                    tk.Label(right_stats_box, text=right_stats, fg='#F2B705', bg='#000000', justify='left', anchor='e', font=('Segoe UI', right_stats_fs, 'bold')).pack(fill='both', expand=True)

                    # scale factor for adi/bdi ball text
                    ADIBDI_TEXT_SCALE = 1.4
                    def draw_ball(parent, relx, rely, value, fill_color, outline_color, text_color, diameter=None):
                        try:
                            txt = str(value if value is not None else '').strip()
                            if txt == '':
                                return
                            if diameter is None:
                                base_d = max(30, min(64, int(stats_lane_h * 0.40), int(w * 0.125)))
                                text_ref_d = max(30, int(base_d * 1.4))
                                d = max(30, int(base_d * 1.5))
                            else:
                                d = int(diameter)
                                text_ref_d = int(diameter)
                            c = tk.Canvas(parent, width=d, height=d, bg='#000000', bd=0, highlightthickness=0)
                            c.place(relx=relx, rely=rely, anchor='center')
                            c.create_oval(1, 1, d - 1, d - 1, fill=fill_color, outline=outline_color, width=1)
                            if len(txt) <= 2:
                                tfs = max(12, int(text_ref_d * 0.44))
                            elif len(txt) <= 3:
                                tfs = max(11, int(text_ref_d * 0.37))
                            else:
                                tfs = max(9, int(text_ref_d * 0.30))
                            # apply base reduction then requested scale (140%)
                            tfs = max(8, int(tfs * 0.70 * ADIBDI_TEXT_SCALE))
                            c.create_text(d/2, d/2, text=txt, fill=text_color, font=('Segoe UI', tfs, 'bold'))
                        except Exception:
                            pass

                    adi_text = str(adi if adi is not None else '').strip()
                    bdi_text = str(bdi if bdi is not None else '').strip()
                    has_adi = adi_text != ''
                    has_bdi = bdi_text != ''
                    ball_y = 0.60

                    if has_adi and has_bdi:
                        draw_ball(stats_area, 0.33, ball_y, adi_text, '#F4F4F4', '#DCDCDC', '#111111')
                        draw_ball(stats_area, 0.67, ball_y, bdi_text, '#F2B705', '#DAA200', '#171717')
                    elif has_adi:
                        draw_ball(stats_area, 0.33, ball_y, adi_text, '#F4F4F4', '#DCDCDC', '#111111')
                    elif has_bdi:
                        draw_ball(stats_area, 0.67, ball_y, bdi_text, '#F2B705', '#DAA200', '#171717')

                    # increase Lượt cơ font size by ~130%
                    lco_fs = _fit_font_size(str(lco), 'Segoe UI', 'bold', _fs(0.20, 31, 78), int(w * 0.24), min_size=11)
                    tk.Label(stats_area, text=f'{lco}', fg='#FF9E9E', bg='#000000', font=('Segoe UI', lco_fs, 'bold')).place(relx=0.5, rely=0.53, anchor='center')

                    # Single subtle gray strip across the full bottom width, touching cell bottom.
                    footer = tk.Frame(base, bg='#2B3138', bd=0, relief='flat', highlightthickness=0)
                    footer.place(relx=0.0, rely=0.86, relwidth=1.0, relheight=0.14)
                    footer_fs_base = _fit_font_size(table_text, 'Segoe UI', 'bold', _fs(0.114, 16, 44), int(w * 0.45), min_size=9)
                    footer_fs = max(7, int(footer_fs_base * 0.40))
                    tran_fs_base = _fit_font_size(tran_text, 'Segoe UI', 'bold', _fs(0.092, 12, 34), int(w * 0.45), min_size=8)
                    tran_fs = max(7, int(tran_fs_base * 0.70))
                    tk.Label(footer, text=table_text, fg='#ECEFF3', bg='#2B3138', anchor='w', font=('Segoe UI', footer_fs, 'bold')).place(relx=0.03, rely=0.5, anchor='w')
                    logo_img = get_footer_logo_tkimg(int(w * 0.22), int(h * 0.105))
                    if logo_img is not None:
                        logo_lbl = tk.Label(footer, image=logo_img, bg='#2B3138', bd=0, highlightthickness=0)
                        logo_lbl.image = logo_img
                        logo_lbl.place(relx=0.5, rely=0.5, anchor='center')
                    tk.Label(footer, text=tran_text, fg='#DCE2E8', bg='#2B3138', anchor='e', font=('Segoe UI', tran_fs, 'bold')).place(relx=0.97, rely=0.5, anchor='e')
                except Exception:
                    pass

            def render_cell_local(idx):
                cell = preview.cell_meta[idx]
                frame = preview.cells[idx]
                w, h = cell_size(frame)
                if w <= 1 or h <= 1:
                    return

                cell_type = cell.get('type') if isinstance(cell, dict) else None
                if cell_type != 'logo_playlist':
                    try:
                        st_map = getattr(preview, '_logo_playlist_states', None)
                        if isinstance(st_map, dict):
                            st_map.pop(idx, None)
                    except Exception:
                        pass
                signature = None
                vmix_url = ''
                cache_snapshot = None
                value_signature = None
                row_signature = None
                if cell_type == 'image':
                    signature = ('image', cell.get('value'), cell.get('image_mode', 'fit'), w, h)
                elif cell_type == 'logo_playlist':
                    logos_signature = tuple(normalize_logo_paths(cell.get('value')))
                    signature = (
                        'logo_playlist',
                        logos_signature,
                        normalize_logo_effect(cell.get('logo_effect')),
                        round(normalize_logo_interval(cell.get('logo_interval'), 4.0), 2),
                        w,
                        h,
                        int(time.monotonic() * 15),
                    )
                elif cell_type == 'vmix':
                    value_signature = cell.get('value')
                    row_signature = resolve_row_index(cell)
                    vmix_url = resolve_vmix_url(cell)
                    if vmix_url:
                        schedule_vmix_fetch(vmix_url, force=False)
                    cache_snapshot = get_cache_snapshot(vmix_url) if vmix_url else None
                    cache_state = 'none'
                    data_fingerprint = None
                    table_hint = ''
                    if cache_snapshot:
                        data_payload = cache_snapshot.get('data') if isinstance(cache_snapshot, dict) else None
                        data_fingerprint = build_vmix_fingerprint(data_payload)
                        try:
                            table_hint = resolve_table_label(cell, vmix_data=data_payload)
                        except Exception:
                            table_hint = ''
                        if data_fingerprint:
                            cache_state = 'ready'
                        elif cache_snapshot.get('loading'):
                            cache_state = 'loading'
                        elif cache_snapshot.get('error'):
                            cache_state = 'error'
                    signature = ('vmix', value_signature, row_signature, vmix_url, cache_state, data_fingerprint, table_hint, w, h)
                else:
                    signature = ('empty', w, h)

                if preview._render_signatures[idx] == signature:
                    return
                preview._render_signatures[idx] = signature

                if cell_type == 'logo_playlist':
                    try:
                        render_logo_playlist_cell(idx, frame, cell, w, h)
                    except Exception:
                        pass
                    return

                for child in list(frame.winfo_children()):
                    try:
                        child.destroy()
                    except Exception:
                        try:
                            child.destroy()
                        except Exception:
                            pass

                if cell_type == 'image':
                    path = cell.get('value')
                    try:
                        if not (Image and ImageTk):
                            raise RuntimeError('Pillow chưa được cài đặt')
                        with Image.open(path) as base_img:
                            img = base_img.copy()
                        mode = cell.get('image_mode', 'fit')
                        iw, ih = img.size
                        try:
                            resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                        except Exception:
                            resample = Image.BICUBIC
                        if mode == 'center':
                            scale = min(1.0, min(w/iw, h/ih))
                            nw = int(iw*scale)
                            nh = int(ih*scale)
                            img2 = img.resize((nw, nh), resample)
                            tkimg = ImageTk.PhotoImage(img2)
                            lbl = tk.Label(frame, image=tkimg, bg='#000')
                            lbl.image = tkimg
                            lbl.place(x=(w-nw)//2, y=(h-nh)//2, width=nw, height=nh)
                        elif mode == 'fit':
                            scale = min(w/iw, h/ih)
                            nw = max(1, int(iw*scale))
                            nh = max(1, int(ih*scale))
                            img2 = img.resize((nw, nh), resample)
                            tkimg = ImageTk.PhotoImage(img2)
                            lbl = tk.Label(frame, image=tkimg, bg='#000')
                            lbl.image = tkimg
                            lbl.place(x=(w-nw)//2, y=(h-nh)//2, width=nw, height=nh)
                        else:
                            scale = max(w/iw, h/ih)
                            nw = max(1, int(iw*scale))
                            nh = max(1, int(ih*scale))
                            img2 = img.resize((nw, nh), resample)
                            left = (nw - w)//2
                            top = (nh - h)//2
                            img3 = img2.crop((left, top, left + w, top + h))
                            tkimg = ImageTk.PhotoImage(img3)
                            lbl = tk.Label(frame, image=tkimg, bg='#000')
                            lbl.image = tkimg
                            lbl.place(x=0, y=0, width=w, height=h)
                    except Exception:
                        tk.Label(frame, text='Image error', fg='white', bg='red').pack(expand=True)
                elif cell_type == 'vmix':
                    try:
                        if not vmix_url:
                            return

                        schedule_vmix_fetch(vmix_url, force=False)
                        cache_snapshot = cache_snapshot or get_cache_snapshot(vmix_url) or {}
                        data = cache_snapshot.get('data') if isinstance(cache_snapshot, dict) else None
                        if data:
                            table_label = resolve_table_label(cell, vmix_data=data)
                            render_vmix_cell_like_sample(frame, data, table_label, w, h)
                    except Exception:
                        pass
                else:
                    ph = tk.Frame(frame, bg='#020202')
                    ph.place(x=0, y=0, relwidth=1.0, relheight=1.0)
                    tk.Frame(ph, bg='#1B1B1B').place(relx=0.14, rely=0.50, relwidth=0.72, relheight=0.003)
                    ph_fs = max(11, min(28, int(min(w, h) * 0.085)))
                    tk.Label(ph, text='POSTER / LOGO', fg='#7A6740', bg='#020202', font=('Segoe UI', ph_fs, 'bold')).place(relx=0.5, rely=0.50, anchor='center')

            def config_cell(idx):
                dlg = tk.Toplevel(preview)
                dlg.title(f'Config ô {idx+1}')
                try:
                    dlg.transient(preview); dlg.grab_set()
                except Exception:
                    pass

                def set_vmix_from_row():
                    sel = simpledialog.askinteger('Chọn hàng', f'Nhập số hàng (0..{max(0, len(getattr(self, "match_rows", []))-1)})')
                    if sel is None:
                        return
                    try:
                        preview.cell_meta[idx] = {'type': 'vmix', 'value': int(sel), 'image_ref': None}
                        preview._render_signatures[idx] = None
                        render_cell_local(idx)
                        schedule_vmix_fetch(resolve_vmix_url(preview.cell_meta[idx]), force=True)
                        try:
                            self._update_last_preview_meta(preview)
                        except Exception:
                            pass
                    except Exception:
                        messagebox.showerror('Lỗi', 'Chọn không hợp lệ')

                def set_vmix_from_url():
                    url = simpledialog.askstring('vMix URL', 'Nhập URL vMix')
                    if not url:
                        return
                    preview.cell_meta[idx] = {'type': 'vmix', 'value': url, 'image_ref': None}
                    preview._render_signatures[idx] = None
                    render_cell_local(idx)
                    schedule_vmix_fetch(resolve_vmix_url(preview.cell_meta[idx]), force=True)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                def set_image():
                    path = filedialog.askopenfilename(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                    if not path:
                        return
                    mode = ask_image_mode_listbox(dlg, initial='fit')
                    if not mode:
                        return
                    preview.cell_meta[idx] = {'type': 'image', 'value': path, 'image_ref': None, 'image_mode': mode}
                    preview._render_signatures[idx] = None
                    render_cell_local(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                def set_logo_playlist():
                    paths = filedialog.askopenfilenames(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                    if not paths:
                        return
                    effect = ask_logo_effect_listbox(dlg, initial='cut')
                    if not effect:
                        return
                    interval = simpledialog.askfloat(
                        'Thời gian đổi logo',
                        'Nhập thời gian đổi logo (giây):',
                        minvalue=0.5,
                        maxvalue=120.0,
                        initialvalue=4.0,
                        parent=dlg,
                    )
                    if interval is None:
                        return
                    preview.cell_meta[idx] = {
                        'type': 'logo_playlist',
                        'value': [str(p) for p in paths],
                        'image_ref': None,
                        'image_mode': 'fit',
                        'logo_effect': str(effect),
                        'logo_interval': float(interval),
                    }
                    preview._render_signatures[idx] = None
                    render_cell_local(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                def clear_cell():
                    preview.cell_meta[idx] = {'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'}
                    preview._render_signatures[idx] = None
                    render_cell_local(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                def set_footer_logo():
                    path = filedialog.askopenfilename(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                    if not path:
                        return
                    try:
                        self._preview_footer_logo_path = path
                        preview.footer_logo_path = path
                    except Exception:
                        pass
                    try:
                        preview._footer_logo_cache = {'path': None, 'size': None, 'tkimg': None}
                    except Exception:
                        pass
                    try:
                        preview._render_signatures = [None for _ in range(9)]
                    except Exception:
                        pass
                    for ii in range(9):
                        try:
                            render_cell_local(ii)
                        except Exception:
                            pass
                    try:
                        self._auto_save_state()
                    except Exception:
                        pass

                def clear_footer_logo():
                    try:
                        self._preview_footer_logo_path = ''
                        preview.footer_logo_path = ''
                    except Exception:
                        pass
                    try:
                        preview._footer_logo_cache = {'path': None, 'size': None, 'tkimg': None}
                    except Exception:
                        pass
                    try:
                        preview._render_signatures = [None for _ in range(9)]
                    except Exception:
                        pass
                    for ii in range(9):
                        try:
                            render_cell_local(ii)
                        except Exception:
                            pass
                    try:
                        self._auto_save_state()
                    except Exception:
                        pass

                tk.Button(dlg, text='Chọn từ bảng', command=set_vmix_from_row, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Nhập URL vMix', command=set_vmix_from_url, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Tải ảnh', command=set_image, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Tải nhiều logo + hiệu ứng', command=set_logo_playlist, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Tải Logo footer', command=set_footer_logo, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Xóa Logo footer', command=clear_footer_logo, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Xóa', command=clear_cell, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Đóng', command=dlg.destroy, width=30).pack(padx=8, pady=8)

            # Always create a fixed 3x3 grid, never resize or change number of cells
            for r in range(3):
                preview.grid_rowconfigure(r, weight=1, minsize=1, uniform='preview_row')
            for c in range(3):
                preview.grid_columnconfigure(c, weight=1, minsize=1, uniform='preview_col')

            for r in range(3):
                for c in range(3):
                    i = r*3 + c
                    f = tk.Frame(
                        preview,
                        bg='#000000',
                        bd=0,
                        relief='flat',
                        highlightthickness=1,
                        highlightbackground='#B7B7B7',
                        highlightcolor='#B7B7B7'
                    )
                    f.grid(row=r, column=c, sticky='nsew', padx=2, pady=2)
                    try:
                        f.grid_propagate(False)
                        f.pack_propagate(False)
                    except Exception:
                        pass
                    preview.cells.append(f)
                    # Bind to <Configure> to auto-scale content inside cell
                    f.bind('<Configure>', lambda e, ii=i: render_cell_local(ii))
                    # Initial render
                    render_cell_local(i)

            ctrl = tk.Frame(preview, bg='#111')
            ctrl.grid(row=3, column=0, columnspan=3, sticky='ew')
            auto_var = tk.IntVar(value=1)
            try:
                tk.Checkbutton(ctrl, text='Auto-refresh', variable=auto_var, bg='#111', fg='white').pack(side='left', padx=8)
            except Exception:
                pass
            btn_close = tk.Button(ctrl, text='Đóng preview', command=on_close, bg='#FF6F00')
            btn_close.pack(side='right', padx=8, pady=6)

            def map_default_seq():
                # Map table rows sequentially while always leaving center cell (index 4) empty.
                try:
                    n = max(0, len(getattr(self, 'match_rows', [])))
                    cell_order = [0, 1, 2, 3, 5, 6, 7, 8]
                    for i in range(9):
                        preview.cell_meta[i] = {'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'}
                    for table_idx, cell_idx in enumerate(cell_order):
                        if table_idx >= n:
                            break
                        preview.cell_meta[cell_idx] = {'type': 'vmix', 'value': table_idx, 'image_ref': None, 'image_mode': 'fit'}
                    preview._render_signatures = [None for _ in range(9)]
                    for i in range(9):
                        try:
                            render_cell_local(i)
                            if preview.cell_meta[i].get('type') == 'vmix':
                                schedule_vmix_fetch(resolve_vmix_url(preview.cell_meta[i]), force=True)
                        except Exception:
                            pass
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass
                except Exception:
                    pass

            try:
                tk.Button(ctrl, text='Mặc định 1→1', command=map_default_seq, bg='#228BE6', fg='white').pack(side='left', padx=8)
            except Exception:
                pass

            def choose_footer_logo_ctrl():
                path = filedialog.askopenfilename(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                if not path:
                    return
                try:
                    self._preview_footer_logo_path = path
                    preview.footer_logo_path = path
                except Exception:
                    pass
                try:
                    preview._footer_logo_cache = {'path': None, 'size': None, 'tkimg': None}
                except Exception:
                    pass
                preview._render_signatures = [None for _ in range(9)]
                for ii in range(9):
                    try:
                        render_cell_local(ii)
                    except Exception:
                        pass
                try:
                    self._auto_save_state()
                except Exception:
                    pass

            def clear_footer_logo_ctrl():
                try:
                    self._preview_footer_logo_path = ''
                    preview.footer_logo_path = ''
                except Exception:
                    pass
                try:
                    preview._footer_logo_cache = {'path': None, 'size': None, 'tkimg': None}
                except Exception:
                    pass
                preview._render_signatures = [None for _ in range(9)]
                for ii in range(9):
                    try:
                        render_cell_local(ii)
                    except Exception:
                        pass
                try:
                    self._auto_save_state()
                except Exception:
                    pass

            try:
                tk.Button(ctrl, text='Tải Logo footer', command=choose_footer_logo_ctrl, bg='#1565C0', fg='white').pack(side='left', padx=8)
                tk.Button(ctrl, text='Xóa Logo footer', command=clear_footer_logo_ctrl, bg='#455A64', fg='white').pack(side='left', padx=8)
            except Exception:
                pass

            def refresh_loop():
                if not preview_alive():
                    return
                try:
                    if auto_var.get():
                        for i, meta in enumerate(preview.cell_meta):
                            if meta and meta.get('type') == 'vmix':
                                try:
                                    vmix_url = resolve_vmix_url(meta)
                                    if vmix_url:
                                        schedule_vmix_fetch(vmix_url, force=False)
                                    render_cell_local(i)
                                except Exception:
                                    pass
                    preview.after(1200, refresh_loop)
                except Exception:
                    try:
                        if preview_alive():
                            preview.after(1200, refresh_loop)
                    except Exception:
                        pass

            def playlist_loop():
                if not preview_alive():
                    return
                try:
                    has_playlist = False
                    for i, meta in enumerate(getattr(preview, 'cell_meta', []) or []):
                        try:
                            if isinstance(meta, dict) and meta.get('type') == 'logo_playlist':
                                has_playlist = True
                                render_cell_local(i)
                        except Exception:
                            pass
                    delay = 70 if has_playlist else 450
                    preview.after(delay, playlist_loop)
                except Exception:
                    try:
                        if preview_alive():
                            preview.after(120, playlist_loop)
                    except Exception:
                        pass

            # start refresh loop
            try:
                preview.after(1200, refresh_loop)
            except Exception:
                pass
            try:
                preview.after(120, playlist_loop)
            except Exception:
                pass

            return preview
        except Exception as ex:
            try:
                print('open_preview_all error', ex)
            except Exception:
                pass
            return None
        try:
            if hasattr(self, 'tengiai_var') and s.get('tengiai'):
                self.tengiai_var.set(s['tengiai'])
            if hasattr(self, 'thoigian_var') and s.get('thoigian'):
                self.thoigian_var.set(s['thoigian'])
            if hasattr(self, 'diadiem_var') and s.get('diadiem'):
                self.diadiem_var.set(s['diadiem'])
            if hasattr(self, 'chuchay_var') and s.get('chuchay'):
                self.chuchay_var.set(s['chuchay'])
            if hasattr(self, 'diemso_text') and s.get('diemso'):
                self.diemso_text.delete('1.0', 'end')
                self.diemso_text.insert('1.0', s['diemso'])
            if hasattr(self, 'url_var') and s.get('sheet_url'):
                self.url_var.set(s['sheet_url'])
            if hasattr(self, 'creds_path') and s.get('creds_path'):
                self.creds_path = s['creds_path']
                if hasattr(self, 'creds_label'):
                    self.creds_label.config(text=os.path.basename(self.creds_path), fg='#00FF00')
            # Khôi phục các trường HBSF
            if s.get('hbsf_url') and hasattr(self, 'hbsf_url_var'):
                self.hbsf_url_var.set(s['hbsf_url'])
            if s.get('event_id') is not None and hasattr(self, 'event_id_var'):
                self.event_id_var.set(str(s.get('event_id', '')))
            if s.get('round_type') and hasattr(self, 'round_type_var'):
                self.round_type_var.set(s['round_type'])
            if hasattr(self, 'ban_var') and s.get('ban'):
                self.ban_var.set(s['ban'])
            # Khôi phục bảng
            if hasattr(self, 'match_rows') and s.get('table'):
                for i, rowdata in enumerate(s['table']):
                    if i < len(self.match_rows):
                        for j in range(min(6, len(rowdata))):
                            self.match_rows[i][j].config(state='normal')
                            self.match_rows[i][j].delete(0, 'end')
                            self.match_rows[i][j].insert(0, rowdata[j])
                            if j in [2,3]:
                                self.match_rows[i][j].config(state='readonly')
        except Exception as ex:
            print(f"[WARN] Không thể gán lại trạng thái UI: {ex}")

    def _toggle_fullscreen(self):
        try:
            self._is_fullscreen = not getattr(self, '_is_fullscreen', False)
            try:
                self.attributes('-fullscreen', self._is_fullscreen)
            except Exception:
                if self._is_fullscreen:
                    try:
                        self.state('zoomed')
                    except Exception:
                        pass
                else:
                    try:
                        self.state('normal')
                    except Exception:
                        pass
        except Exception:
            pass

    @staticmethod
    def _safe_widget_get(widget, default='', text_widget=False):
        """Gọi .get() an toàn, trả về default nếu widget đã bị destroy."""
        try:
            if text_widget:
                return widget.get('1.0', 'end-1c')
            return widget.get()
        except Exception:
            return default

    def _auto_save_state(self):
        """Lưu trạng thái UI (sheet_url, credentials, ban, table rows, header fields) vào file pickle."""
        # Không autosave khi đang khôi phục trạng thái
        if getattr(self, '_restoring_state', False):
            return
        # If a restore is in-progress (not yet committed), skip autosave to avoid overwriting
        if not getattr(self, '_restore_committed', True):
            return
        state = {}
        state['tengiai'] = self._safe_widget_get(self.tengiai_var) if hasattr(self, 'tengiai_var') else ''
        state['thoigian'] = self._safe_widget_get(self.thoigian_var) if hasattr(self, 'thoigian_var') else ''
        state['diadiem'] = self._safe_widget_get(self.diadiem_var) if hasattr(self, 'diadiem_var') else ''
        state['chuchay'] = self._safe_widget_get(self.chuchay_var) if hasattr(self, 'chuchay_var') else ''
        state['diemso'] = self._safe_widget_get(self.diemso_text, default='', text_widget=True) if hasattr(self, 'diemso_text') else ''
        state['sheet_url'] = self._safe_widget_get(self.url_var) if hasattr(self, 'url_var') else ''
        state['creds_path'] = self.creds_path if hasattr(self, 'creds_path') else None
        # HBSF fields
        state['hbsf_url'] = self._safe_widget_get(self.hbsf_url_var, default='https://hbsf.com.vn') if hasattr(self, 'hbsf_url_var') else 'https://hbsf.com.vn'
        state['event_id'] = self._safe_widget_get(self.event_id_var) if hasattr(self, 'event_id_var') else ''
        state['round_type'] = self._safe_widget_get(self.round_type_var, default='Vòng Loại') if hasattr(self, 'round_type_var') else 'Vòng Loại'
        # Luôn lưu số bàn, kể cả khi là 0 hoặc None
        try:
            state['ban'] = int(self.ban_var.get()) if hasattr(self, 'ban_var') and self.ban_var is not None else 0
        except Exception:
            state['ban'] = 0
        # serialize current table (first 6 columns) for each row
        table = []
        for r in getattr(self, 'match_rows', []):
            rowvals = []
            for j in range(6):
                w = r[j]
                if hasattr(w, 'get'):
                    rowvals.append(self._safe_widget_get(w))
                else:
                    rowvals.append('')
            table.append(rowvals)
        state['table'] = table
        # serialize preview configuration (use last-known meta if preview window closed)
        preview = getattr(self, '_preview_window', None)
        serial_preview = None
        if preview is not None and getattr(preview, 'winfo_exists', lambda: False)():
            pm = getattr(preview, 'cell_meta', None)
            if pm:
                serial_preview = []
                for cell in pm:
                    serial_preview.append({
                        'type': cell.get('type'),
                        'value': cell.get('value'),
                        'image_mode': cell.get('image_mode', 'fit'),
                        'logo_effect': cell.get('logo_effect', 'cut'),
                        'logo_interval': cell.get('logo_interval', 4.0),
                    })
                # update last-known snapshot
                self._last_preview_meta = serial_preview
        else:
            # use previously stored preview meta if available
            serial_preview = getattr(self, '_last_preview_meta', None)
        state['preview'] = serial_preview
        state['preview_user_configured'] = bool(getattr(self, '_preview_meta_user_configured', False))
        state['preview_footer_logo'] = str(getattr(self, '_preview_footer_logo_path', '') or '')
        state['_schema_version'] = int(getattr(self, '_state_schema_version', 2) or 2)
        try:
            import platform
            state['_machine_name'] = str(platform.node() or '')
        except Exception:
            state['_machine_name'] = ''
        # write to file
        should_write = True
        try:
            if os.path.exists(self._auto_state_path):
                import pickle as _pickle
                with open(self._auto_state_path, 'rb') as _f:
                    try:
                        existing = _pickle.load(_f)
                        if existing == state:
                            should_write = False
                    except Exception:
                        should_write = True
        except Exception:
            should_write = True

        if should_write:
            try:
                parent_dir = os.path.dirname(self._auto_state_path)
                if parent_dir:
                    os.makedirs(parent_dir, exist_ok=True)
            except Exception:
                pass
            try:
                with open(self._auto_state_path, 'wb') as f:
                    import pickle
                    pickle.dump(state, f)
            except Exception:
                pass
        # if suspended due to recent restore, skip this save cycle
        import time
        if getattr(self, '_autosave_suspended_until', 0) and time.time() < self._autosave_suspended_until:
            # reschedule quickly to recheck
            self.after(1000, self._periodic_autosave)
        # ...existing code...
        try:
            # reschedule
            self.after(5000, self._periodic_autosave)
        except Exception:
            pass

    def manual_load_state(self):
        """Allow user to pick a ui_state.pkl file and apply it immediately (debug helper)."""
        try:
            from tkinter import filedialog, messagebox
            import pickle
            # try default candidates first (no prompt)
            candidates = []
            if getattr(self, '_auto_state_path', None):
                candidates.append(self._auto_state_path)
            candidates.append(os.path.join(os.getcwd(), 'ui_state.pkl'))
            candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ui_state.pkl'))
            try:
                import sys
                exe_dir = os.path.dirname(sys.executable)
                candidates.append(os.path.join(exe_dir, 'ui_state.pkl'))
            except Exception:
                pass
            p = None
            s = None
            for cand in candidates:
                try:
                    if cand and os.path.exists(cand):
                        p = cand
                        with open(p, 'rb') as f:
                            s = pickle.load(f)
                        break
                except Exception:
                    p = None
                    s = None
                    continue
            if s is None:
                # fallback to asking user
                p = filedialog.askopenfilename(title='Chọn ui_state.pkl', filetypes=[('Pickle files','*.pkl'), ('All files','*.*')])
                if not p:
                    return
                try:
                    with open(p, 'rb') as f:
                        s = pickle.load(f)
                except Exception as ex:
                    try:
                        messagebox.showerror('Lỗi', f'Không đọc được file: {ex}')
                    except Exception:
                        pass
                    return
            try:
                self._restored_state = s
                self._auto_restore_state_to_ui()
                try:
                    messagebox.showinfo('Load trạng thái', f'Đã load {p}. Khôi phục {len(s.get("table",[]))} hàng; preview_present={bool(s.get("preview"))}')
                except Exception:
                    pass
            except Exception as ex:
                try:
                    messagebox.showerror('Lỗi', f'Không áp dụng được trạng thái: {ex}')
                except Exception:
                    self._restoring_state = True
        except Exception:
            pass

    def _on_close_and_save_state(self):
        self._auto_save_state()
        # Xóa file tạm nếu có
        # import os, sys đã ở đầu file
        if hasattr(sys, '_MEIPASS'):
            exe_dir = os.path.dirname(sys.executable)
        else:
            exe_dir = os.path.dirname(os.path.abspath(__file__))
        temp_path = os.path.join(exe_dir, 'vmix_input1_temp.csv')
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception as ex:
            print(f"[WARN] Không thể xóa file tạm: {ex}")
        self.destroy()

    def propagate_master_score(self, value: str):
        """Gán giá trị `value` vào cột Điểm của tất cả các hàng hiện có."""
        try:
            if not hasattr(self, 'match_rows'):
                return
            for row in getattr(self, 'match_rows', []) or []:
                try:
                    diem_entry = row[4]
                    if hasattr(diem_entry, 'delete') and hasattr(diem_entry, 'insert'):
                        diem_entry.delete(0, 'end')
                        diem_entry.insert(0, value)
                except Exception:
                    continue
        except Exception:
            pass
        try:
            # persist the change immediately
            self._auto_save_state()
        except Exception:
            pass

    def _update_last_preview_meta(self, preview_window):
        """Store a serializable snapshot of preview cell_meta on the main window.
        This allows restoring preview even when the preview Toplevel is closed.
        """
        try:
            if preview_window is None:
                return
            pm = getattr(preview_window, 'cell_meta', None)
            if not pm:
                return
            serial_preview = []
            for cell in pm:
                try:
                    serial_preview.append({
                        'type': cell.get('type'),
                        'value': cell.get('value'),
                        'image_mode': cell.get('image_mode', 'fit'),
                        'logo_effect': cell.get('logo_effect', 'cut'),
                        'logo_interval': cell.get('logo_interval', 4.0),
                    })
                except Exception:
                    serial_preview.append(None)
            self._last_preview_meta = serial_preview
        except Exception:
            pass

    def _apply_preview_footer_logo_to_open_window(self):
        try:
            p = getattr(self, '_preview_window', None)
            if p is None or not getattr(p, 'winfo_exists', lambda: False)():
                return
            try:
                p.footer_logo_path = str(getattr(self, '_preview_footer_logo_path', '') or '')
            except Exception:
                p.footer_logo_path = ''
            try:
                p._footer_logo_cache = {'path': None, 'size': None, 'tkimg': None}
            except Exception:
                pass
            try:
                p._render_signatures = [None for _ in range(9)]
            except Exception:
                pass
            for frame in getattr(p, 'cells', []) or []:
                try:
                    frame.event_generate('<Configure>')
                except Exception:
                    pass
        except Exception:
            pass

    def save_preview_now(self):
        """Force-save the current preview configuration into the autosave file."""
        p = getattr(self, '_preview_window', None)
        if p is not None and getattr(p, 'winfo_exists', lambda: False)():
            try:
                self._update_last_preview_meta(p)
            except Exception:
                pass
        # always run autosave (which will include last-known preview)
        self._auto_save_state()
        try:
            self.status_var.set('Đã lưu preview')
        except Exception:
            pass
        try:
            self._append_debug_log('Toolbar: Save preview now')
        except Exception:
            pass

    def open_preview_mapping_dialog(self):
        """Open a dialog to configure preview cell mappings without opening the preview window."""
        try:
            from tkinter import simpledialog, filedialog, messagebox
            dlg = tk.Toplevel(self)
            dlg.title('Cấu hình Preview')
            try:
                dlg.transient(self); dlg.grab_set()
            except Exception:
                pass

            # load current serial preview meta or default
            serial = getattr(self, '_last_preview_meta', None)
            if not serial or not isinstance(serial, list) or len(serial) < 9:
                serial = [None for _ in range(9)]

            footer_logo_var = tk.StringVar(value=os.path.basename(str(getattr(self, '_preview_footer_logo_path', '') or '')) or '(Chưa chọn logo)')

            def choose_footer_logo():
                try:
                    path = filedialog.askopenfilename(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                except Exception:
                    path = ''
                if not path:
                    return
                try:
                    self._preview_footer_logo_path = path
                    footer_logo_var.set(os.path.basename(path))
                except Exception:
                    pass
                try:
                    self._apply_preview_footer_logo_to_open_window()
                except Exception:
                    pass
                try:
                    self._auto_save_state()
                except Exception:
                    pass

            def clear_footer_logo():
                try:
                    self._preview_footer_logo_path = ''
                    footer_logo_var.set('(Chưa chọn logo)')
                except Exception:
                    pass
                try:
                    self._apply_preview_footer_logo_to_open_window()
                except Exception:
                    pass
                try:
                    self._auto_save_state()
                except Exception:
                    pass

            cells_frames = []

            def pretty(cell):
                try:
                    if not cell:
                        return 'Empty'
                    t = cell.get('type')
                    v = cell.get('value')
                    if t == 'vmix':
                        ban_name = cell.get('ban_name') if isinstance(cell, dict) else None
                        if ban_name:
                            return f'Bàn {ban_name}'
                        return f'vMix: {v}'
                    if t == 'image':
                        return f'Image: {os.path.basename(v) if v else v}'
                    if t == 'logo_playlist':
                        logos = v if isinstance(v, (list, tuple)) else []
                        effect = str(cell.get('logo_effect') or 'cut')
                        interval = cell.get('logo_interval', 4.0)
                        try:
                            interval_txt = f'{float(interval):.1f}s'
                        except Exception:
                            interval_txt = '4.0s'
                        return f'Playlist: {len(logos)} logo ({effect}, {interval_txt})'
                    return str(cell)
                except Exception:
                    return 'Invalid'

            for i in range(9):
                fr = tk.Frame(dlg, bd=1, relief='ridge', padx=6, pady=4)
                fr.grid(row=i//3, column=i%3, padx=6, pady=6, sticky='nsew')
                lbl = tk.Label(fr, text=f'Ô {i+1}', font=('Arial', 11, 'bold'))
                lbl.pack(anchor='w')
                cur = tk.Label(fr, text=pretty(serial[i]), wraplength=180, justify='left')
                cur.pack(anchor='w')

                def make_set_vmix_row(ii, cur_label):
                    def _():
                        try:
                            sel = simpledialog.askinteger('Chọn hàng', f'Nhập số hàng (0..{max(0, len(getattr(self, "match_rows", []))-1)})')
                            if sel is None:
                                return
                            serial[ii] = {'type': 'vmix', 'value': int(sel)}
                            cur_label.config(text=pretty(serial[ii]))
                            # apply to open preview window if exists
                            try:
                                self.preview_set_cell(ii, 'vmix', int(sel))
                            except Exception:
                                pass
                        except Exception:
                            messagebox.showerror('Lỗi', 'Chọn không hợp lệ')
                    return _

                def make_set_vmix_url(ii, cur_label):
                    def _():
                        url = simpledialog.askstring('vMix URL', 'Nhập URL vMix')
                        if not url:
                            return
                        serial[ii] = {'type': 'vmix', 'value': url}
                        cur_label.config(text=pretty(serial[ii]))
                        try:
                            self.preview_set_cell(ii, 'vmix', url)
                        except Exception:
                            pass
                    return _

                def make_set_image(ii, cur_label):
                    def _():
                        path = filedialog.askopenfilename(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                        if not path:
                            return
                        mode = ask_image_mode_listbox(dlg, initial='fit')
                        if not mode:
                            return
                        serial[ii] = {'type': 'image', 'value': path, 'image_mode': mode}
                        cur_label.config(text=pretty(serial[ii]))
                        try:
                            self.preview_set_cell(ii, 'image', path, image_mode=mode)
                        except Exception:
                            pass
                    return _

                def make_set_logo_playlist(ii, cur_label):
                    def _():
                        paths = filedialog.askopenfilenames(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                        if not paths:
                            return
                        effect = ask_logo_effect_listbox(dlg, initial='cut')
                        if not effect:
                            return
                        interval = simpledialog.askfloat(
                            'Thời gian đổi logo',
                            'Nhập thời gian đổi logo (giây):',
                            minvalue=0.5,
                            maxvalue=120.0,
                            initialvalue=4.0,
                            parent=dlg,
                        )
                        if interval is None:
                            return
                        serial[ii] = {
                            'type': 'logo_playlist',
                            'value': [str(p) for p in paths],
                            'image_mode': 'fit',
                            'logo_effect': str(effect),
                            'logo_interval': float(interval),
                        }
                        cur_label.config(text=pretty(serial[ii]))
                        try:
                            self.preview_set_cell(
                                ii,
                                'logo_playlist',
                                list(paths),
                                logo_effect=effect,
                                logo_interval=interval,
                            )
                        except Exception:
                            pass
                    return _

                def make_clear(ii, cur_label):
                    def _():
                        serial[ii] = None
                        cur_label.config(text=pretty(serial[ii]))
                        try:
                            self.preview_set_cell(ii, 'clear', None)
                        except Exception:
                            pass
                    return _

                def make_show_ban_listbox(ii, cur_label):
                    def _():
                        ban_names = []
                        if hasattr(self, 'match_rows'):
                            for row in self.match_rows:
                                ban_val = row[1].get().strip() if len(row) > 1 else ''
                                if ban_val and ban_val not in ban_names:
                                    ban_names.append(ban_val)
                        ban_win = tk.Toplevel(dlg)
                        ban_win.title('Chọn Bàn')
                        tk.Label(ban_win, text='Chọn Bàn:', font=('Segoe UI', 13, 'bold')).pack(padx=12, pady=8)
                        lb_ban = tk.Listbox(ban_win, height=min(8, max(1, len(ban_names))), font=('Segoe UI', 13), width=24)
                        for name in ban_names:
                            lb_ban.insert('end', name)
                        lb_ban.pack(padx=12, pady=8)
                        def on_ban_select(event=None):
                            sel = lb_ban.curselection()
                            if sel:
                                selected_ban = lb_ban.get(sel[0])
                                vmix_addr = ''
                                for row in self.match_rows:
                                    ban_val = row[1].get().strip() if len(row) > 1 else ''
                                    if ban_val == selected_ban and len(row) > 5:
                                        vmix_addr = row[5].get().strip()
                                        break
                                serial[ii] = {'type': 'vmix', 'value': str(vmix_addr), 'ban_name': selected_ban}
                                cur_label.config(text=f'Bàn {selected_ban}')
                                try:
                                    self.preview_set_cell(ii, 'vmix', str(vmix_addr))
                                except Exception:
                                    pass
                                ban_win.destroy()
                        lb_ban.bind('<<ListboxSelect>>', on_ban_select)
                    return _

                btn_row = tk.Button(fr, text='Chọn Bàn', command=make_show_ban_listbox(i, cur), width=18)
                btn_row.pack(pady=2)
                btn_url = tk.Button(fr, text='Nhập URL vMix', command=make_set_vmix_url(i, cur), width=18)
                btn_url.pack(pady=2)
                btn_img = tk.Button(fr, text='Tải ảnh', command=make_set_image(i, cur), width=18)
                btn_img.pack(pady=2)
                btn_playlist = tk.Button(fr, text='Nhiều logo + FX', command=make_set_logo_playlist(i, cur), width=18)
                btn_playlist.pack(pady=2)
                btn_clr = tk.Button(fr, text='Xóa', command=make_clear(i, cur), width=18)
                btn_clr.pack(pady=2)

                cells_frames.append((fr, cur))

            logo_row = tk.Frame(dlg)
            logo_row.grid(row=3, column=0, columnspan=3, pady=(4, 2), padx=10, sticky='ew')
            tk.Label(logo_row, text='Logo footer:', font=('Segoe UI', 10, 'bold')).pack(side='left', padx=(0, 8))
            tk.Label(logo_row, textvariable=footer_logo_var, fg='#0D47A1').pack(side='left', padx=(0, 8))
            tk.Button(logo_row, text='Tải Logo', command=choose_footer_logo, width=12).pack(side='left', padx=4)
            tk.Button(logo_row, text='Xóa Logo', command=clear_footer_logo, width=12).pack(side='left', padx=4)

            def apply_and_close():
                try:
                    # normalize serial length
                    sp = [None]*9
                    for k in range(9):
                        sp[k] = serial[k] if k < len(serial) else None
                    self._last_preview_meta = sp
                    # đánh dấu người dùng đã explicitly cấu hình
                    self._preview_meta_user_configured = True
                    try:
                        self._auto_save_state()
                    except Exception:
                        pass
                except Exception:
                    pass
                try:
                    dlg.destroy()
                except Exception:
                    pass

            btn_apply = tk.Button(dlg, text='Áp dụng và lưu', command=apply_and_close, bg='#00C853', fg='white')
            btn_apply.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky='w')
            btn_close = tk.Button(dlg, text='Đóng', command=lambda: dlg.destroy())
            btn_close.grid(row=4, column=2, pady=10, padx=10, sticky='e')
            return dlg
        except Exception:
            return None

    def _restore_preview_later(self, preview_data):
        """Attempt to open preview and apply `preview_data` (used when immediate restore failed)."""
        try:
            try:
                dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                import datetime
                with open(dbg, 'a', encoding='utf-8') as df:
                    df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE_PREVIEW retry invoked\n")
            except Exception:
                pass
            p = self.preview_open()
            if p is None:
                return
            # rebuild cell_meta similar to main restore
            new_meta = []
            for cell in preview_data:
                try:
                    if cell is None:
                        new_meta.append({'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'})
                    else:
                        new_meta.append({
                            'type': cell.get('type'),
                            'value': cell.get('value'),
                            'image_ref': None,
                            'image_mode': cell.get('image_mode', 'fit'),
                            'logo_effect': cell.get('logo_effect', 'cut'),
                            'logo_interval': cell.get('logo_interval', 4.0),
                        })
                except Exception:
                    new_meta.append({'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'})
            try:
                p.cell_meta = new_meta
            except Exception:
                pass
            try:
                for i, f in enumerate(getattr(p, 'cells', []) or []):
                    try:
                        f.event_generate('<Configure>')
                    except Exception:
                        try:
                            f.update_idletasks()
                        except Exception:
                            pass
            except Exception:
                pass
            try:
                # record snapshot
                self._update_last_preview_meta(p)
            except Exception:
                pass
        except Exception:
            pass

    def _apply_restored_table(self, table, attempt=0, max_attempts=12):
        """Apply `table` values into `self.match_rows` when the widgets are ready.
        Retries a few times if `match_rows` hasn't been populated yet.
        """
        try:
            desired = len(table or [])
            cur = len(getattr(self, 'match_rows', []) or [])
            if cur < desired and attempt < max_attempts:
                # schedule a retry after a short delay to let populate_table finish
                try:
                    self.after(150, lambda: self._apply_restored_table(table, attempt+1, max_attempts))
                except Exception:
                    pass
                return
            # apply values to available rows
            for i, rowvals in enumerate(table or []):
                if i >= len(getattr(self, 'match_rows', []) or []):
                    break
                for j in range(min(len(rowvals), 6)):
                    try:
                        widget = self.match_rows[i][j]
                        if hasattr(widget, 'config') and hasattr(widget, 'delete'):
                            try:
                                widget.config(state='normal')
                            except Exception:
                                pass
                            try:
                                widget.delete(0, 'end')
                                widget.insert(0, rowvals[j])
                            except Exception:
                                pass
                            try:
                                if j in (2, 3):
                                    widget.config(state='readonly')
                            except Exception:
                                pass
                    except Exception:
                        continue
        except Exception:
            pass

    # Toolbar handlers (explicit, with logging) to ensure button clicks produce observable effects
    def _append_debug_log(self, text):
        try:
            import sys, os, datetime
            exe_dir = os.path.dirname(sys.executable) if hasattr(sys, '_MEIPASS') else os.path.dirname(os.path.abspath(__file__))
            p = os.path.join(exe_dir, 'vmix_debug.log')
            with open(p, 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.datetime.now().isoformat()}] {text}\n")
        except Exception:
            pass

    def toolbar_load(self):
        try:
            self._append_debug_log('Toolbar: Tải bảng pressed')
            # Tải từ file đã lưu
            self.load_table_from_csv()
            try:
                self.status_var.set('Tải bảng: OK')
            except Exception:
                pass
            try:
                from tkinter import messagebox
                messagebox.showinfo('Tải bảng', 'Đã tải lại bảng dữ liệu.')
            except Exception:
                pass
        except Exception as ex:
            try:
                self._append_debug_log(f'Toolbar load error: {ex}')
                messagebox.showerror('Lỗi', f'Không tải được bảng: {ex}')
            except Exception:
                pass

    def toolbar_save(self):
        try:
            self._append_debug_log('Toolbar: Lưu bảng pressed')
            # always persist UI state locally
            self._auto_save_state()
            try:
                self.status_var.set('Lưu trạng thái: OK')
            except Exception:
                pass

            # Prepare a safe AA..AK batch preview for Google Sheets (dry-run)
            try:
                from tkinter import messagebox
                import re, os
                sheet_url = self.url_var.get().strip() if hasattr(self, 'url_var') else ''
                m = re.search(r"/spreadsheets/d/([\w-]+)", sheet_url) if sheet_url else None
                spreadsheet_id = m.group(1) if m else None
                creds_path = self.creds_path if hasattr(self, 'creds_path') else None
                batch = []
                if spreadsheet_id and creds_path and os.path.exists(creds_path) and hasattr(self, 'match_rows'):
                    try:
                        gs = GSheetClient(spreadsheet_id, creds_path)
                        rows = gs.read_table('Kết quả!A1:Z2000') or []
                    except Exception:
                        rows = []

                    # helper to normalize/compare Trận values
                    def normalize_tran(val):
                        try:
                            s = str(val)
                            import re
                            m2 = re.search(r"(\d+)", s)
                            if m2:
                                return int(m2.group(1))
                            return s.strip().lower()
                        except Exception:
                            return val

                    headers = list(rows[0].keys()) if rows else []

                    for i, widgets in enumerate(self.match_rows):
                        try:
                            tran_val = widgets[0].get().strip() if hasattr(widgets[0], 'get') else ''
                            if not tran_val:
                                continue
                            key = normalize_tran(tran_val)
                            found_idx = None
                            for ri, r in enumerate(rows):
                                sheet_tran = r.get(headers[0], '') if headers else ''
                                # attempt to find a matching row by scanning common 'Trận' like keys
                                # check all keys for a numeric match
                                for k in (headers or []):
                                    st = r.get(k, '')
                                    if normalize_tran(st) == key:
                                        found_idx = ri
                                        break
                                if found_idx is not None:
                                    break
                            if found_idx is None:
                                continue
                            rownum = found_idx + 2
                            # collect vMix-derived values from UI row
                            ten_a = widgets[2].get() if hasattr(widgets[2], 'get') else ''
                            ten_b = widgets[3].get() if hasattr(widgets[3], 'get') else ''
                            diem_a = widgets[4].get() if hasattr(widgets[4], 'get') else ''
                            diem_b = ''
                            lco = ''
                            hr1a = ''
                            hr2a = ''
                            hr1b = ''
                            hr2b = ''
                            # try to parse additional fields from earlier per-row state if available
                            try:
                                # some rows may have been filled from vMix earlier and stored in hidden attrs
                                vmix_url = widgets[5].get() if hasattr(widgets[5], 'get') else ''
                            except Exception:
                                vmix_url = ''
                            # compute AVGA/AVGB
                            def to_float(v):
                                try:
                                    if v is None or v == '':
                                        return None
                                    return float(str(v).replace(',', '.'))
                                except Exception:
                                    return None
                            a = to_float(diem_a)
                            c = to_float(lco)
                            avga = (round(a/c, 3) if a is not None and c and c != 0 else '')
                            avgb = ''

                            # build contiguous AA..AK (11 cols) values: TenA, TenB, DiemA, DiemB, Lco, HR1A, HR2A, HR1B, HR2B, AVGA, AVGB
                            values = [ten_a, ten_b, diem_a, diem_b, lco, hr1a, hr2a, hr1b, hr2b, (f"{avga:.3f}" if avga != '' else ''), (f"{avgb:.3f}" if avgb != '' else '')]
                            cell_range = f'Kết quả!AA{rownum}:AK{rownum}'
                            batch.append({'range': cell_range, 'values': [values]})
                        except Exception:
                            continue

                # Log preview of PREWRITE items
                try:
                    dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                except Exception:
                    dbg = os.path.join(os.getcwd(), 'vmix_debug.log')
                try:
                    import datetime
                    with open(dbg, 'a', encoding='utf-8') as df:
                        df.write(f"[{datetime.datetime.now().isoformat()}] PREWRITE preview items={len(batch)}\n")
                        for it in batch:
                            df.write(f"[{datetime.datetime.now().isoformat()}] PREWRITE_ITEM range={it.get('range')} values={it.get('values')}\n")
                except Exception:
                    pass

                # Ask user to confirm performing the real safe write
                if batch:
                    do_write = messagebox.askyesno('Xác nhận ghi', f'Có {len(batch)} hàng chuẩn bị ghi vào cột AA..AK. Thực hiện ghi lên Google Sheets?')
                    if do_write:
                        try:
                            gs.batch_update_safe(batch)
                            try:
                                import datetime
                                with open(dbg, 'a', encoding='utf-8') as df:
                                    df.write(f"[{datetime.datetime.now().isoformat()}] WRITE_OK batch items={len(batch)}\n")
                            except Exception:
                                pass
                            messagebox.showinfo('Ghi xong', 'Đã ghi an toàn các ô đã chọn lên Google Sheets.')
                        except Exception as ex:
                            try:
                                import datetime
                                with open(dbg, 'a', encoding='utf-8') as df:
                                    df.write(f"[{datetime.datetime.now().isoformat()}] WRITE_FAIL {ex}\n")
                            except Exception:
                                pass
                            messagebox.showerror('Lỗi ghi', f'Không ghi được: {ex}')
                else:
                    messagebox.showinfo('Preview', 'Không có hàng nào để ghi (hoặc chưa cấu hình Google Sheets).')
            except Exception as ex:
                try:
                    self._append_debug_log(f'Toolbar save preview error: {ex}')
                except Exception:
                    pass
        except Exception as ex:
            try:
                self._append_debug_log(f'Toolbar save error: {ex}')
                messagebox.showerror('Lỗi', f'Không lưu được: {ex}')
            except Exception:
                pass

    def toolbar_exit(self):
        try:
            self._append_debug_log('Toolbar: Thoát pressed')
            self._on_close_and_save_state()
        except Exception:
            try:
                self.destroy()
            except Exception:
                pass

    def create_widgets(self):

        # (Top toolbar removed to keep titlebar accessible; controls moved to bottom)



        # Header frame for global info (multi-row layout)
        header_frame = tk.Frame(self, bg='#1A2233', highlightbackground='#222C3A', highlightthickness=2)
        header_frame.pack(fill='x', pady=8, padx=(2, 6))
        # Icon (if any)
        col_offset = 0
        if hasattr(self, 'icon_img'):
            tk.Label(header_frame, image=self.icon_img, bg='#222831').grid(row=0, column=0, rowspan=2, padx=8, pady=2, sticky='w')
            col_offset = 1
        # Configure columns: giữ nhãn đủ chỗ hiển thị, ưu tiên giãn cho ô nhập liệu
        max_col = 12 + col_offset
        for i in range(max_col):
            header_frame.grid_columnconfigure(i, weight=0)
        for i in [col_offset+1, col_offset+2, col_offset+3, col_offset+5, col_offset+6, col_offset+8, col_offset+9, col_offset+10, col_offset+11]:
            header_frame.grid_columnconfigure(i, weight=1)
        for i in [col_offset+0, col_offset+4, col_offset+7]:
            header_frame.grid_columnconfigure(i, weight=0, minsize=152)
        # Row 0: TÊN GIẢI - THỜI GIAN - ĐIỂM SỐ
        tk.Label(header_frame, text='Tên giải:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=0, column=col_offset+0, sticky='e', padx=(4,2), pady=4)
        self.tengiai_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.tengiai_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369', width=60).grid(row=0, column=col_offset+1, columnspan=3, sticky='ew', padx=(0,6), pady=4)
        tk.Label(header_frame, text='Thời gian:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=0, column=col_offset+4, sticky='e', padx=(4,2), pady=4)
        self.thoigian_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.thoigian_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369', width=40).grid(row=0, column=col_offset+5, columnspan=2, sticky='ew', padx=(0,6), pady=4)
        tk.Label(header_frame, text='Điểm số:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=0, column=col_offset+7, sticky='e', padx=(4,2), pady=4)
        self.diemso_var = tk.StringVar()
        self.diemso_text = tk.Text(header_frame, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369', height=2, width=24, wrap='word')
        self.diemso_text.grid(row=0, column=col_offset+8, columnspan=2, sticky='ew', padx=(0,6), pady=4)
        if hasattr(self, 'diemso_var'):
            self.diemso_text.insert('1.0', self.diemso_var.get())
        # Cập nhật tất cả ô điểm số khi bấm Enter hoặc Ctrl+Enter
        def update_all_scores(event=None):
            value = self.diemso_text.get('1.0', 'end-1c')
            try:
                # delegate to propagate_master_score to ensure consistent behavior
                self.propagate_master_score(value)
            except Exception:
                try:
                    if hasattr(self, 'match_rows'):
                        for row in self.match_rows:
                            diem_entry = row[4]
                            if isinstance(diem_entry, tk.Entry):
                                diem_entry.delete(0, 'end')
                                diem_entry.insert(0, value)
                except Exception:
                    pass
        self.diemso_text.bind('<Return>', lambda e: (update_all_scores(), 'break'))
        self.diemso_text.bind('<Control-Return>', lambda e: (update_all_scores(), 'break'))
        # also apply when user tabs out or clicks away
        try:
            self.diemso_text.bind('<FocusOut>', lambda e: (self.propagate_master_score(self.diemso_text.get('1.0','end-1c'))))
        except Exception:
            pass
        # Row 1: ĐỊA ĐIỂM
        tk.Label(header_frame, text='Địa điểm:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=1, column=col_offset+0, sticky='e', padx=(4,2), pady=4)
        self.diadiem_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.diadiem_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369', width=120).grid(row=1, column=col_offset+1, columnspan=9, sticky='ew', padx=(0,6), pady=4)
        # Row 2: CHỮ CHẠY
        tk.Label(header_frame, text='Chữ chạy:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=2, column=col_offset+0, sticky='e', padx=(4,2), pady=4)
        self.chuchay_var = tk.StringVar()
        self.chuchay_text = tk.Text(header_frame, font=('Segoe UI', 16), relief='groove', bd=3, bg='#232B3E', fg='white', height=4, wrap='word', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369', width=160)
        self.chuchay_text.grid(row=2, column=col_offset+1, columnspan=9, sticky='ew', padx=(0,6), pady=4)
        # Đồng bộ Text widget với StringVar cũ (get/set)
        def update_chuchay_var(event=None):
            self.chuchay_var.set(self.chuchay_text.get('1.0', 'end-1c'))
        self.chuchay_text.bind('<KeyRelease>', update_chuchay_var)
        # Khi set var, update Text widget
        def set_chuchay_text(val):
            self.chuchay_text.delete('1.0', 'end')
            self.chuchay_text.insert('1.0', val)
        self.chuchay_var.trace_add('write', lambda *a: set_chuchay_text(self.chuchay_var.get()))
        # Nếu có giá trị ban đầu, set luôn
        if self.chuchay_var.get():
            set_chuchay_text(self.chuchay_var.get())

        # ── HBSF: Lấy thông tin trận đấu từ website ─────────────────────────────
        hbsf_frame = tk.Frame(self, bg='#232B3E')
        hbsf_frame.pack(fill='x', pady=2)

        # status_var phải khởi tạo TRƯỚC khi tạo button vì load_hbsf_matches dùng ngay
        self.status_var = tk.StringVar(value='')

        # Label + Entry: HBSF URL
        tk.Label(hbsf_frame, text='HBSF URL:', fg='#FFD369', bg='#232B3E', font=('Segoe UI', 13, 'bold')).pack(side='left', padx=(10, 2))
        self.hbsf_url_var = tk.StringVar(value=getattr(self, '_saved_hbsf_url', 'https://hbsf.com.vn'))
        hbsf_url_entry = tk.Entry(hbsf_frame, textvariable=self.hbsf_url_var, width=22, font=('Segoe UI', 13), relief='groove', bd=2, bg='#1A2233', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369')
        hbsf_url_entry.pack(side='left', padx=(0, 8))

        # Label + Entry: Event ID
        tk.Label(hbsf_frame, text='Event ID:', fg='#FFD369', bg='#232B3E', font=('Segoe UI', 13, 'bold')).pack(side='left', padx=(0, 2))
        self.event_id_var = tk.StringVar(value=getattr(self, '_saved_event_id', ''))
        event_id_entry = tk.Entry(hbsf_frame, textvariable=self.event_id_var, width=8, font=('Segoe UI', 13), relief='groove', bd=2, bg='#1A2233', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369')
        event_id_entry.pack(side='left', padx=(0, 8))

        # Label + OptionMenu: Loại vòng
        tk.Label(hbsf_frame, text='Loại vòng:', fg='#FFD369', bg='#232B3E', font=('Segoe UI', 13, 'bold')).pack(side='left', padx=(0, 2))
        self.round_type_var = tk.StringVar(value=getattr(self, '_saved_round_type', 'Vòng Loại'))
        round_type_menu = tk.OptionMenu(hbsf_frame, self.round_type_var, 'Vòng Loại', 'Vòng Chính Thức')
        round_type_menu.config(bg='#1A2233', fg='white', activebackground='#FFD369', activeforeground='#1A2233', font=('Segoe UI', 13), relief='groove', bd=2, highlightthickness=0)
        round_type_menu['menu'].config(bg='#1A2233', fg='white', activebackground='#FFD369', activeforeground='#1A2233', font=('Segoe UI', 13))
        round_type_menu.pack(side='left', padx=(0, 8))

        # Button: Lấy thông tin trận đấu
        tk.Button(hbsf_frame, text='Lấy thông tin trận đấu', command=self.load_hbsf_matches,
                  bg='#FFD369', fg='#1A2233', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2,
                  activebackground='#FFE082', activeforeground='#1A2233').pack(side='left', padx=5)
        event_id_entry.bind('<Return>', lambda e: self.load_hbsf_matches())

        # Button: Xem Log
        tk.Button(hbsf_frame, text='Xem Log', command=self.show_log_popup,
                  bg='#546E7A', fg='white', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2,
                  activebackground='#78909C', activeforeground='white').pack(side='left', padx=5)

        # Các biến Google Sheet (giữ lại để tương thích với code lưu trạng thái)
        self.url_var = tk.StringVar()
        self.creds_label = tk.Label(hbsf_frame, text='', fg='#232B3E', bg='#232B3E')  # hidden

        # ── (CODE GOOGLE SHEET CŨ — ĐÃ COMMENT OUT, KHÔNG XÓA) ────────────────
        # url_frame = tk.Frame(self, bg='#232B3E')
        # url_frame.pack(fill='x', pady=2)
        # tk.Label(url_frame, text='Link Google Sheet:', fg='#FFD369', bg='#232B3E', font=('Segoe UI', 14, 'bold')).pack(side='left', padx=10)
        # self.url_var = tk.StringVar()
        # url_entry = tk.Entry(url_frame, textvariable=self.url_var, width=40, font=('Segoe UI', 14), relief='groove', bd=2, bg='#1A2233', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369')
        # url_entry.pack(side='left', padx=5)
        # url_entry.bind('<Return>', lambda e: self.reload_matches())
        # def open_gsheet_link(event=None):
        #     import webbrowser
        #     url = self.url_var.get().strip()
        #     if url:
        #         webbrowser.open(url)
        # url_entry.bind('<Double-Button-1>', open_gsheet_link)
        # tk.Button(url_frame, text='Chọn credentials', command=self.select_credentials, bg='#FFD369', fg='#1A2233', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#FFE082', activeforeground='#1A2233').pack(side='left', padx=5)
        # self.creds_label = tk.Label(url_frame, text='(Chưa chọn credentials)', fg='#FFD369', bg='#232B3E', font=('Segoe UI', 11, 'italic'))
        # self.creds_label.pack(side='left', padx=5)
        # try:
        #     cp = getattr(self, 'creds_path', None)
        #     if cp and os.path.exists(cp):
        #         self.creds_label.config(text=os.path.basename(cp), fg='#00FF00')
        # except Exception:
        #     pass
        # self.status_var = tk.StringVar(value='')
        # tk.Button(url_frame, text='Tải dữ liệu', command=self.reload_matches, bg='#FFD369', fg='#1A2233', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#FFE082', activeforeground='#1A2233').pack(side='left', padx=10)
        # def preview_sheet():
        #     import tkinter as tk
        #     import json
        #     win = tk.Toplevel(self)
        #     win.base_title = 'Preview dữ liệu Google Sheet'
        #     win.title(win.base_title)
        #     try:
        #         self.animate_title(win, win.base_title)
        #     except Exception:
        #         try:
        #             win.title(win.base_title)
        #         except Exception:
        #             pass
        #     win.geometry('800x400')
        #     text = tk.Text(win, font=('Consolas', 11), wrap='none')
        #     text.pack(fill='both', expand=True)
        #     if self.sheet_rows:
        #         try:
        #             pretty = json.dumps(self.sheet_rows, ensure_ascii=False, indent=2)
        #         except Exception:
        #             pretty = str(self.sheet_rows)
        #         text.insert('1.0', pretty)
        #     else:
        #         text.insert('1.0', 'Không có dữ liệu Google Sheet!')
        #     text.config(state='disabled')
        # tk.Button(url_frame, text='Preview Sheet', command=preview_sheet, bg='#FFD369', fg='#1A2233', font=('Segoe UI', 11), relief='groove', bd=2, activebackground='#FFE082', activeforeground='#1A2233').pack(side='left', padx=5)
        # ── (KẾT THÚC CODE GOOGLE SHEET CŨ) ───────────────────────────────────

        # Nút Lưu bảng / Tải bảng (giữ lại)
        try:
            self.save_btn = tk.Button(hbsf_frame, text='Lưu bảng', command=self.save_table_to_csv, bg='#00C853', fg='#1A2233', font=('Segoe UI', 11, 'bold'), relief='groove', bd=2, activebackground='#B9F6CA', activeforeground='#1A2233')
            self.load_btn = tk.Button(hbsf_frame, text='Tải bảng', command=self.load_table_from_csv, bg='#2962FF', fg='#FFD369', font=('Segoe UI', 11, 'bold'), relief='groove', bd=2, activebackground='#82B1FF', activeforeground='#FFD369')
            self.save_btn.pack(side='left', padx=8)
            self.load_btn.pack(side='left', padx=4)
        except Exception:
            pass

        # Nút preview tổng hợp bảng điểm
        def open_preview_all():
            """Open a fullscreen 3x3 preview grid. Each cell can load a vMix scoreboard (from a match row or URL) or an image file.

            The preview stays full-screen with 9 boxes; unused boxes show posters/logos (image) or remain blank.
            """
            import requests
            import xml.etree.ElementTree as ET
            from tkinter import filedialog, simpledialog
            try:
                from PIL import Image, ImageTk
            except Exception:
                Image = ImageTk = None

            # Create preview window
            preview = tk.Toplevel(self)
            # expose preview window on self for automated tests
            try:
                self._preview_window = preview
            except Exception:
                pass
            preview.base_title = 'Preview 3x3 Bảng điểm'
            try:
                preview.title(preview.base_title)
            except Exception:
                pass
            # Try fullscreen; fallback to maximized
            try:
                preview.attributes('-fullscreen', True)
            except Exception:
                try:
                    preview.state('zoomed')
                except Exception:
                    pass

            # grid of 3x3 always
            preview.cells = []
            preview.cell_meta = [{'type': None, 'value': None, 'image_ref': None, 'img_mode': 'fit'} for _ in range(9)]

            def render_cell(idx):
                frame = preview.cells[idx]
                # clear
                for w in frame.winfo_children():
                    w.destroy()
                meta = preview.cell_meta[idx]

                # Helper to get cell pixel size (fallback to one-third screen)
                def cell_size(frame):
                    try:
                        fw = max(10, frame.winfo_width())
                        fh = max(10, frame.winfo_height())
                    except Exception:
                        fw = preview.winfo_screenwidth() // 3
                        fh = preview.winfo_screenheight() // 3
                    return fw, fh

                if meta['type'] == 'vmix' and meta['value']:
                    # meta['value'] is either a match row index or a URL string
                    try:
                        if isinstance(meta['value'], int):
                            rowidx = meta['value']
                            if rowidx < len(self.match_rows):
                                row = self.match_rows[rowidx]
                                vmix_url = row[5].get().strip() if hasattr(row[5], 'get') else ''
                            else:
                                vmix_url = ''
                        else:
                            vmix_url = str(meta['value'])
                        if vmix_url:
                            resp = requests.get(f'{vmix_url}/API/', timeout=3)
                            root = ET.fromstring(resp.text)
                            input1 = root.find(".//input[@number='1']")
                            def get_field(name):
                                if input1 is not None:
                                    for text in input1.findall('text'):
                                        if text.attrib.get('name') == name:
                                            return text.text or ''
                                return ''
                            ten_a = get_field('TenA.Text') or get_field('TenA') or ''
                            ten_b = get_field('TenB.Text') or get_field('TenB') or ''
                            tran = get_field('Tran.Text') or ''
                            diem_a = get_field('DiemA.Text') or get_field('DiemA') or ''
                            diem_b = get_field('DiemB.Text') or get_field('DiemB') or ''
                            avg_a = get_field('AvgA.Text') or get_field('AVGA') or ''
                            avg_b = get_field('AvgB.Text') or get_field('AVGB') or ''
                            lco = get_field('Lco.Text') or ''
                            # Compose layout similar to scoreboard and scale fonts to fill the cell
                            fw, fh = cell_size(frame)
                            # baseline sizes for a typical cell (approx 800x600)
                            base_w, base_h = preview.winfo_screenwidth() // 3, preview.winfo_screenheight() // 3
                            try:
                                scale = min(float(fw) / float(base_w), float(fh) / float(base_h))
                            except Exception:
                                scale = 1.0
                            def s(x):
                                return max(8, int(x * scale))
                            # Centered header and match label
                            tk.Label(frame, text=ten_a, font=('Arial', s(24), 'bold'), fg='white', bg='#000', wraplength=int(fw*0.9), justify='center').pack(anchor='n', padx=8, pady=(8,0))
                            tk.Label(frame, text=f'TRẬN {tran}', font=('Arial', s(14), 'bold'), fg='#00E5FF', bg='#000').pack()
                            # Use a horizontal frame for scores, center both numbers and scale font to fit
                            score_row = tk.Frame(frame, bg='#000')
                            score_row.pack(fill='both', expand=True)
                            # determine a font size that fits the width roughly
                            try:
                                max_score_width = max(1, fw//2 - 24)
                                score_font_size = s(min(72, max(24, int(max_score_width/ (len(str(diem_a))+1)))))
                            except Exception:
                                score_font_size = s(64)
                            lbl_a = tk.Label(score_row, text=f"{diem_a}", font=('Arial', score_font_size, 'bold'), fg='white', bg='#000')
                            lbl_b = tk.Label(score_row, text=f"{diem_b}", font=('Arial', score_font_size, 'bold'), fg='#FFD600', bg='#000')
                            lbl_a.pack(side='left', expand=True, padx=8)
                            lbl_b.pack(side='right', expand=True, padx=8)
                            # sub info row
                            sub = tk.Frame(frame, bg='#000')
                            sub.pack(fill='x', pady=6)
                            tk.Label(sub, text=f"AVG {avg_a}", fg='white', bg='#000', font=('Arial', s(12))).pack(side='left', padx=6)
                            tk.Label(sub, text=f"Lco {lco}", fg='#FFD369', bg='#000', font=('Arial', s(12))).pack(side='left', padx=6)
                            tk.Label(sub, text=f"AVG {avg_b}", fg='#FFD600', bg='#000', font=('Arial', s(12))).pack(side='right', padx=6)
                        else:
                            return
                    except Exception:
                        return
                elif meta['type'] == 'image' and meta['value']:
                    path = meta['value']
                    mode = meta.get('img_mode', 'fit')
                    if Image and ImageTk:
                        try:
                            img = Image.open(path)
                            fw, fh = cell_size(frame)
                            target_w = max(10, fw - 20)
                            target_h = max(10, fh - 20)
                            if mode == 'fit':
                                # preserve ratio, fit inside
                                img.thumbnail((target_w, target_h), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
                                tkimg = ImageTk.PhotoImage(img)
                                lbl = tk.Label(frame, image=tkimg, bg='#000')
                                lbl.image = tkimg
                                lbl.place(relx=0.5, rely=0.5, anchor='center')
                                preview.cell_meta[idx]['image_ref'] = tkimg
                            elif mode == 'center':
                                # do not scale beyond original size; center it
                                iw, ih = img.size
                                scale = min(1.0, float(target_w) / iw, float(target_h) / ih)
                                if scale < 1.0:
                                    neww = int(iw * scale)
                                    newh = int(ih * scale)
                                    img = img.resize((neww, newh), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
                                tkimg = ImageTk.PhotoImage(img)
                                lbl = tk.Label(frame, image=tkimg, bg='#000')
                                lbl.image = tkimg
                                lbl.place(relx=0.5, rely=0.5, anchor='center')
                                preview.cell_meta[idx]['image_ref'] = tkimg
                            else:
                                # cover: scale to fill and crop center without changing aspect ratio
                                iw, ih = img.size
                                scale = max(float(target_w) / iw, float(target_h) / ih)
                                neww = int(iw * scale)
                                newh = int(ih * scale)
                                img2 = img.resize((neww, newh), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
                                # crop center
                                left = max(0, (neww - target_w)//2)
                                top = max(0, (newh - target_h)//2)
                                right = left + target_w
                                bottom = top + target_h
                                img3 = img2.crop((left, top, right, bottom))
                                tkimg = ImageTk.PhotoImage(img3)
                                lbl = tk.Label(frame, image=tkimg, bg='#000')
                                lbl.image = tkimg
                                lbl.place(relx=0.5, rely=0.5, anchor='center')
                                preview.cell_meta[idx]['image_ref'] = tkimg
                        except Exception as ex:
                            tk.Label(frame, text=f'Image error:\n{ex}', fg='red', bg='#000').pack(padx=6, pady=6)
                    else:
                        tk.Label(frame, text='Pillow not installed\nCannot display image', fg='white', bg='#000').pack(padx=6, pady=6)
                else:
                    # empty placeholder / poster
                    tk.Label(frame, text='Poster / Logo', fg='#FFD369', bg='#000', font=('Arial', 20, 'bold')).pack(expand=True)

            def config_cell(idx):
                # Small dialog to choose source for this cell
                dlg = tk.Toplevel(preview)
                dlg.title(f'Config ô {idx+1}')
                try:
                    dlg.transient(preview)
                    dlg.grab_set()
                except Exception:
                    pass

                def set_vmix_from_row():
                    # choose match row
                    options = []
                    for i, r in enumerate(self.match_rows):
                        try:
                            ban = r[1].get() if hasattr(r[1], 'get') else f'Bàn {i+1}'
                            a = r[2].get() if hasattr(r[2], 'get') else ''
                            b = r[3].get() if hasattr(r[3], 'get') else ''
                            options.append((i, f"{ban}: {a} vs {b}"))
                        except Exception:
                            options.append((i, f'Row {i}'))
                    if not options:
                        messagebox.showinfo('Info', 'Không có hàng nào trong bảng để chọn')
                        return
                    sel = simpledialog.askinteger('Chọn hàng', f'Nhập số hàng (0..{len(options)-1}) để dùng làm nguồn vMix')
                    if sel is None:
                        return
                    try:
                        sel = int(sel)
                        preview.cell_meta[idx]['type'] = 'vmix'
                        preview.cell_meta[idx]['value'] = sel
                        render_cell(idx)
                    except Exception:
                        messagebox.showerror('Lỗi', 'Chọn không hợp lệ')

                def set_row_field():
                    # Ask for row index
                    if not getattr(self, 'match_rows', None):
                        messagebox.showinfo('Info', 'Không có hàng nào trong bảng để chọn')
                        return
                    max_idx = len(self.match_rows) - 1
                    row_sel = simpledialog.askinteger('Chọn hàng', f'Nhập số hàng (0..{max_idx}) để dùng làm nguồn')
                    if row_sel is None:
                        return
                    try:
                        row_sel = int(row_sel)
                    except Exception:
                        messagebox.showerror('Lỗi', 'Chọn không hợp lệ')
                        return
                    # Ask for field via a small dropdown dialog (better UX)
                    fields = [('match', 'Số trận'), ('table', 'Số bàn'), ('name_a', 'Tên VĐV A'), ('name_b', 'Tên VĐV B'), ('score', 'Điểm số'), ('vmix', 'Địa chỉ vMix')]
                    fld_keys = [k for k, _ in fields]
                    fld_labels = {k:lbl for k,lbl in fields}
                    pick = tk.Toplevel(dlg)
                    pick.title('Chọn trường hiển thị')
                    try:
                        pick.transient(dlg)
                        pick.grab_set()
                    except Exception:
                        pass
                    sel_var = tk.StringVar(value=fld_keys[0])
                    tk.Label(pick, text=f'Chọn trường cho hàng {row_sel}:').pack(padx=8, pady=(8,4))
                    opts = []
                    for k in fld_keys:
                        opts.append(f"{k} - {fld_labels.get(k,'')}")
                    # Use OptionMenu showing key+label, store key by splitting on ' - '
                    opt_var = tk.StringVar(value=opts[0])
                    tk.OptionMenu(pick, opt_var, *opts).pack(padx=8, pady=6)
                    def on_ok():
                        choice = opt_var.get()
                        key = choice.split(' - ')[0].strip()
                        preview.cell_meta[idx]['type'] = 'row_field'
                        preview.cell_meta[idx]['value'] = (row_sel, key)
                        try:
                            self._update_last_preview_meta(preview)
                        except Exception:
                            pass
                        render_cell(idx)
                        try:
                            pick.destroy()
                        except Exception:
                            pass
                    def on_cancel():
                        try:
                            pick.destroy()
                        except Exception:
                            pass
                    btnf = tk.Frame(pick)
                    btnf.pack(pady=8)
                    tk.Button(btnf, text='OK', command=on_ok, width=10).pack(side='left', padx=6)
                    tk.Button(btnf, text='Hủy', command=on_cancel, width=10).pack(side='left', padx=6)

                def set_vmix_from_url():
                    url = simpledialog.askstring('vMix URL', 'Nhập URL vMix (ví dụ http://192.168.1.2:8088)')
                    if not url:
                        return
                    preview.cell_meta[idx]['type'] = 'vmix'
                    preview.cell_meta[idx]['value'] = url
                    render_cell(idx)

                def set_image():
                    path = filedialog.askopenfilename(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp'), ('All', '*.*')])
                    if not path:
                        return
                    mode = ask_image_mode_listbox(dlg, initial='fit')
                    if not mode:
                        return
                    preview.cell_meta[idx]['type'] = 'image'
                    preview.cell_meta[idx]['value'] = path
                    preview.cell_meta[idx]['image_mode'] = mode
                    render_cell(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                def clear_cell():
                    preview.cell_meta[idx] = {'type': None, 'value': None, 'image_ref': None}
                    render_cell(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                tk.Button(dlg, text='Chọn từ bảng (hàng - full vMix)', command=set_vmix_from_row, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Chọn trường cụ thể từ hàng', command=set_row_field, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Nhập URL vMix', command=set_vmix_from_url, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Tải ảnh (Logo/Poster)', command=set_image, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Xóa', command=clear_cell, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Đóng', command=dlg.destroy, width=30).pack(padx=8, pady=8)

            # Create 3x3 frames (fixed grid; cells should not be reflowed)
            for r in range(3):
                preview.grid_rowconfigure(r, weight=1)
                for c in range(3):
                    preview.grid_columnconfigure(c, weight=1)
                    idx = r*3 + c
                    f = tk.Frame(preview, bg='#000', bd=4, relief='ridge')
                    f.grid(row=r, column=c, sticky='nsew', padx=4, pady=4)
                    preview.cells.append(f)
                    # Redraw when cell resizes so content fits exactly
                    f.bind('<Configure>', lambda e, i=idx: render_cell(i))
                    render_cell(idx)

            # Controls at bottom to auto-refresh and exit fullscreen
            ctrl = tk.Frame(preview, bg='#111')
            ctrl.grid(row=3, column=0, columnspan=3, sticky='ew')
            auto_var = tk.IntVar(value=0)

            def render_cell(idx):
                pm = getattr(preview, 'cell_meta', None)
                cell = {}
                try:
                    if isinstance(pm, dict):
                        cell = pm.get(idx, {})
                    elif isinstance(pm, (list, tuple)):
                        cell = pm[idx] if idx < len(pm) else {}
                    else:
                        cell = {}
                except Exception:
                    cell = {}
                frame = preview.cells[idx]
                for w in frame.winfo_children():
                    w.destroy()
                w = frame.winfo_width()
                h = frame.winfo_height()
                if w <= 1 or h <= 1:
                    return
                if cell.get('type') == 'image':
                    path = cell.get('value')
                    try:
                        from PIL import Image, ImageTk
                        img = Image.open(path)
                        mode = cell.get('image_mode', 'fit')
                        iw, ih = img.size
                        if mode == 'center':
                            # scale down only if larger than frame, keep aspect ratio
                            scale = min(1.0, min(w/iw, h/ih))
                            nw = int(iw*scale)
                            nh = int(ih*scale)
                            img2 = img.resize((nw, nh), Image.LANCZOS)
                            tkimg = ImageTk.PhotoImage(img2)
                            lbl = tk.Label(frame, image=tkimg, bg='#000')
                            lbl.image = tkimg
                            lbl.place(x=(w-nw)//2, y=(h-nh)//2, width=nw, height=nh)
                        elif mode == 'fit':
                            scale = min(w/iw, h/ih)
                            nw = max(1, int(iw*scale))
                            nh = max(1, int(ih*scale))
                            img2 = img.resize((nw, nh), Image.LANCZOS)
                            tkimg = ImageTk.PhotoImage(img2)
                            lbl = tk.Label(frame, image=tkimg, bg='#000')
                            lbl.image = tkimg
                            lbl.place(x=(w-nw)//2, y=(h-nh)//2, width=nw, height=nh)
                        else:  # cover
                            scale = max(w/iw, h/ih)
                            nw = max(1, int(iw*scale))
                            nh = max(1, int(ih*scale))
                            img2 = img.resize((nw, nh), Image.LANCZOS)
                            # crop to center
                            left = (nw - w)//2
                            top = (nh - h)//2
                            img3 = img2.crop((left, top, left + w, top + h))
                            tkimg = ImageTk.PhotoImage(img3)
                            lbl = tk.Label(frame, image=tkimg, bg='#000')
                            lbl.image = tkimg
                            lbl.place(x=0, y=0, width=w, height=h)
                    except Exception as e:
                        lbl = tk.Label(frame, text='Image error', fg='white', bg='red')
                        lbl.pack(expand=True)
                elif cell.get('type') == 'row_field':
                    try:
                        row_idx, field = cell.get('value')
                        text = ''
                        # prefer match_rows widgets
                        if getattr(self, 'match_rows', None) and isinstance(row_idx, int) and row_idx < len(self.match_rows):
                            r = self.match_rows[row_idx]
                            fmap = {'match':0, 'table':1, 'name_a':2, 'name_b':3, 'score':4, 'vmix':5}
                            if field in fmap:
                                try:
                                    w = r[fmap[field]]
                                    text = w.get() if hasattr(w, 'get') else ''
                                except Exception:
                                    text = ''
                        else:
                            # fallback to table_rows dict-like
                            if getattr(self, 'table_rows', None) and isinstance(row_idx, int) and row_idx < len(self.table_rows):
                                row = self.table_rows[row_idx]
                                fk = {'match':'Trận','table':'Bàn','name_a':'TenA','name_b':'TenB','score':'Diem','vmix':'Vmix'}
                                key = fk.get(field)
                                if key and isinstance(row, dict):
                                    text = row.get(key,'')
                        # render text
                        from tkinter import font as tkfont
                        font_size = max(10, min(72, int(min(w,h)//12)))
                        fnt = tkfont.Font(family='Arial', size=font_size)
                        lbl = tk.Label(frame, text=text, fg='white', bg='#222', font=fnt, wraplength=int(w*0.98), justify='center')
                        lbl.place(x=0, y=0, width=w, height=h)
                    except Exception as e:
                        lbl = tk.Label(frame, text=f'Error: {e}', fg='white', bg='red')
                        lbl.pack(expand=True)
                elif cell.get('type') == 'vmix':
                    row_idx = cell.get('value')
                    if row_idx is None or row_idx >= len(self.table_rows):
                        lbl = tk.Label(frame, text='No row', fg='white', bg='#111')
                        lbl.pack(expand=True, fill='both')
                        return
                    row = self.table_rows[row_idx]
                    tenA = row.get('TenA','')
                    tenB = row.get('TenB','')
                    dA = row.get('DiemA','')
                    dB = row.get('DiemB','')
                    txt = f"{tenA} {dA} - {dB} {tenB}"
                    # dynamic font sizing to fill cell
                    from tkinter import font as tkfont
                    base = tkfont.Font(family='Arial', size=10)
                    # binary search for font size
                    lo, hi = 8, 240
                    best = 8
                    while lo <= hi:
                        mid = (lo+hi)//2
                        fnt = tkfont.Font(family='Arial', size=mid)
                        tw = fnt.measure(txt)
                        th = fnt.metrics('linespace')
                        if tw <= w*0.95 and th <= h*0.9:
                            best = mid
                            lo = mid + 1
                        else:
                            hi = mid - 1
                    fnt = tkfont.Font(family='Arial', size=best)
                    lbl = tk.Label(frame, text=txt, fg='white', bg='#222', font=fnt, wraplength=int(w*0.98), justify='center')
                    lbl.place(x=0, y=0, width=w, height=h)
                else:
                    lbl = tk.Label(frame, text='Empty', fg='white', bg='#111')
                    lbl.pack(expand=True, fill='both')
        # Outer container for the table area
        try:
            table_outer = tk.Frame(self, bg='#111')
        except Exception:
            table_outer = tk.Frame(self)
        table_outer.pack(fill='both', expand=True, padx=0, pady=10)
        self.table_canvas = tk.Canvas(table_outer, bg='#1A2233', highlightthickness=0)
        self.table_canvas.pack(side='left', fill='both', expand=True)
        v_scroll = tk.Scrollbar(table_outer, orient='vertical', command=self.table_canvas.yview)
        v_scroll.pack(side='right', fill='y')
        self.table_canvas.configure(yscrollcommand=v_scroll.set)
        self.table_frame = tk.Frame(self.table_canvas, bg='#232B3E')
        self.table_window = self.table_canvas.create_window((0,0), window=self.table_frame, anchor='nw')
        # (Đã set tỉ lệ cột phía trên, không override ở đây)
        # Enable mousewheel scrolling for vertical scrollbar
        def _on_mousewheel(event):
            self.table_canvas.yview_scroll(int(-1*(event.delta/120)), 'units')
        self.table_canvas.bind_all('<MouseWheel>', _on_mousewheel)
        # Table header
        header_bg = '#232B3E'
        header_fg = '#FFD369'
        headers = ['Trận', 'BÀN', 'Tên VĐV A', 'Tên VĐV B', 'Điểm số', 'Địa chỉ vMix', 'Kết quả', 'Gửi', 'Sửa']
        # Đặt width rõ ràng cho cột Trận và Số bàn
        for col, text in enumerate(headers):
            label = tk.Label(self.table_frame, text=text, bg=header_bg, fg=header_fg, font=('Arial', 18, 'bold'), relief='raised', bd=2)
            label.grid(row=0, column=col, sticky='nsew', padx=2, pady=2, ipadx=4, ipady=8)
        # store headers for adaptive sizing
        self.table_headers = headers

        # Chia tỉ lệ % chiều ngang cho các cột
        # 8 cột: Trận, Số bàn, Tên VĐV A, Tên VĐV B, Điểm số, Địa chỉ vMix, Kết quả, Gửi
        # Chia lại: Trận 2%, BÀN 10%, các cột còn lại chia lại cho đủ 100%
        # 8 cột: Trận, BÀN, Tên VĐV A, Tên VĐV B, Điểm số, Địa chỉ vMix, Kết quả, Gửi
        # Phân bổ lại: Trận 14%, BÀN 8%, Tên A 15%, Tên B 15%, Điểm số 6%, Địa chỉ vMix 21%, Kết quả 6%, Gửi 6%, Sửa 6%
        # store weights on self so resize handler can compute min pixel sizes
        # Adjusted weights to avoid overly wide name columns on large screens.
        # Order: Trận, BÀN, Tên A, Tên B, Điểm số, Địa chỉ vMix, Kết quả, Gửi, Sửa
        # Rebalanced weights: make 'BÀN' and 'Tên A/B' larger for readability
        # Order: Trận, BÀN, Tên A, Tên B, Điểm số, Địa chỉ vMix, Kết quả, Gửi, Sửa
        # Reduce width for 'BÀN' and 'Điểm số' per user request
        # Order: Trận, BÀN, Tên A, Tên B, Điểm số, Địa chỉ vMix, Kết quả, Gửi, Sửa
        # Adjust weights so 'Điểm số' and 'Địa chỉ vMix' occupy less horizontal space
        # Reduce 'Kết quả', 'Điểm số', and 'Địa chỉ vMix' to roughly half their previous horizontal weight
        # Order: Trận, BÀN, Tên A, Tên B, Điểm số, Địa chỉ vMix, Kết quả, Gửi, Sửa
        # Further reduce 'Kết quả' (index 6) per request
        # Increase Name A/B widths by ~130%
        self.col_weights = [6, 4, 21, 21, 1, 8, 2, 5, 4]
        self._col_total_weight = sum(self.col_weights)
        for col, weight in enumerate(self.col_weights):
            # Use fixed minsize distribution instead of grid weight-driven expansion
            # to have precise control over column widths and avoid overly wide columns.
            self.table_frame.grid_columnconfigure(col, weight=0)
        self.match_rows = []
        # Đảm bảo tạo xong toàn bộ widget trước khi restore
        try:
            import datetime
            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
            with open(dbg, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE_CALL_BEGIN\n")
        except Exception:
            pass
        try:
            self._auto_restore_state_to_ui()
        except Exception:
            pass
        # Debug: log ban_var value after restore
        try:
            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
            import datetime
            with open(dbg, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now().isoformat()}] AFTER_RESTORE ban_var={getattr(self, 'ban_var', None).get() if hasattr(self, 'ban_var') else 'N/A'}\n")
        except Exception:
            pass
        try:
            self.populate_table()
        except Exception:
            pass
        # Debug: log ban_var value after table population
        try:
            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
            import datetime
            with open(dbg, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now().isoformat()}] AFTER_POPULATE_TABLE ban_var={getattr(self, 'ban_var', None).get() if hasattr(self, 'ban_var') else 'N/A'}\n")
        except Exception:
            pass
        try:
            import datetime
            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
            with open(dbg, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE_CALL_END\n")
        except Exception:
            pass
        # schedule periodic autosave after restore to avoid overwriting restored values
        try:
            self.after(5000, self._periodic_autosave)
        except Exception:
            pass
        try:
            self.after(int(getattr(self, 'auto_ketqua_interval_ms', 5000)), self._auto_ketqua_tick)
        except Exception:
            pass

        # Địa chỉ vMix input
        # Địa chỉ vMix đã chuyển xuống từng dòng, không cần ô nhập chung
        # Số bàn input và bộ lọc tìm kiếm chuyển xuống dưới cùng
        self.bottom_frame = tk.Frame(self, bg='#222831')
        # Pack bottom frame after the table; ensure it doesn't get collapsed and stays on top
        self.bottom_frame.pack(fill='x', pady=5, side='bottom')
        try:
            # Prevent automatic shrink/expand so the bar remains visible
            self.bottom_frame.pack_propagate(False)
            self.bottom_frame.configure(height=56)
        except Exception:
            pass
        tk.Label(self.bottom_frame, text='Tổng số bàn:', fg='#FFD369', bg='#222831', font=('Arial', 12, 'bold')).pack(side='left', padx=10)
        ban_spin = tk.Spinbox(self.bottom_frame, from_=1, to=32, textvariable=self.ban_var, width=5, font=('Arial', 12))
        ban_spin.pack(side='left', padx=5)
        # (Đoạn này đã được xử lý ở phía dưới, không cần pack ở đây)
        # Nút cập nhật số bàn: cập nhật lại bảng theo số bàn mới
        tk.Button(self.bottom_frame, text='Cập nhật', command=self.populate_table, bg='#FFD369', fg='#222831', font=('Arial', 11, 'bold')).pack(side='left', padx=10)
        tk.Label(self.bottom_frame, text='Lọc bàn:', fg='#FFD369', bg='#222831', font=('Arial', 11)).pack(side='left', padx=5)
        self.filter_ban_var = tk.StringVar()
        self.filter_ban_entry = tk.Entry(self.bottom_frame, textvariable=self.filter_ban_var, width=8, font=('Arial', 11))
        self.filter_ban_entry.pack(side='left', padx=2)
        tk.Label(self.bottom_frame, text='Tìm VĐV:', fg='#FFD369', bg='#222831', font=('Arial', 11)).pack(side='left', padx=5)
        self.filter_vdv_var = tk.StringVar()
        self.filter_vdv_entry = tk.Entry(self.bottom_frame, textvariable=self.filter_vdv_var, width=16, font=('Arial', 11))
        self.filter_vdv_entry.pack(side='left', padx=2)
        tk.Button(self.bottom_frame, text='Lọc', command=self.populate_table, bg='#FFD369', fg='#222831', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        # Nút 'Preview tổng hợp' đặt gần góc phải, trước dòng trạng thái
        self.preview_all_btn = tk.Button(self.bottom_frame, text='Preview tổng hợp', command=self.open_preview_all, bg='#00B8D4', fg='#fff', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#4DD0E1', activeforeground='#222831')
        self.preview_all_btn.pack(side='right', padx=10)
        try:
            self.preview_map_btn = tk.Button(self.bottom_frame, text='Cấu hình Preview', command=lambda: self.open_preview_mapping_dialog(), bg='#7E57C2', fg='white', font=('Segoe UI', 10, 'bold'))
            self.preview_map_btn.pack(side='right', padx=6)
        except Exception:
            pass
        # Use class method directly so preview window single-instance control is always respected.
        # Thêm dòng báo trạng thái bên phải
        self.status_label = tk.Label(self.bottom_frame, textvariable=self.status_var, fg='#00FF00', bg='#222831', font=('Segoe UI', 12, 'italic'))
        self.status_label.pack(side='right', padx=10)
        # Nút ghi tất cả lên Google Sheets (mở popup preview trước khi ghi)
        # Note: Save/Load buttons are placed in the URL row to ensure visibility.
        self.table_frame.bind('<Configure>', self._on_table_configure)
        self.table_canvas.bind('<Configure>', self._on_canvas_configure)
        # Quick floating access to Save/Load in case bottom bar is not visible on some displays
        try:
            # remove floating quick actions to avoid overlap; Save/Load are in bottom bar
            pass
        except Exception:
            pass
        try:
            # Bring bottom bar to front so buttons are visible
            if hasattr(self, 'bottom_frame'):
                try:
                    self.bottom_frame.lift()
                except Exception:
                    pass
                try:
                    if hasattr(self, 'save_btn'):
                        self.save_btn.lift()
                    if hasattr(self, 'load_btn'):
                        self.load_btn.lift()
                except Exception:
                    pass
        except Exception:
            pass

        # (Save/Load buttons are created in the URL row to avoid header overlap)

    def _on_table_configure(self, event):
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        # Make inner frame width match canvas width if possible
        canvas_width = event.width
        self.table_canvas.itemconfig(self.table_window, width=canvas_width)
        # Adjust per-column minsize so entries resize proportionally to available width
        try:
            if hasattr(self, 'col_weights') and getattr(self, '_col_total_weight', 0) > 0:
                total = float(self._col_total_weight)
                # Account for horizontal padding/gaps introduced by grid `padx` and widget borders.
                # Many grid placements use `padx=2` (left+right => 4px per column). Subtract that
                # total padding from the canvas width before distributing pixels to columns so
                # the sum of minsize values more closely matches the actual available content width.
                try:
                    pad_per_column = 4
                    total_padding = pad_per_column * len(self.col_weights)
                    avail_width = max(0, int(canvas_width) - int(total_padding))
                except Exception:
                    avail_width = int(canvas_width)

                # Adaptive sizing: measure header + cell content pixel widths and derive proportional sizes
                try:
                    import tkinter.font as tkfont
                    header_font = tkfont.Font(family='Arial', size=18, weight='bold')
                    cell_font = tkfont.Font(family='Arial', size=18)
                    measured = []
                    col_count = len(self.col_weights)
                    for col in range(col_count):
                        maxw = 0
                        # measure header
                        try:
                            h = self.table_headers[col] if hasattr(self, 'table_headers') and col < len(self.table_headers) else ''
                            maxw = max(maxw, header_font.measure(h) + 16)
                        except Exception:
                            pass
                        # measure visible rows
                        try:
                            for row in getattr(self, 'match_rows', []):
                                if col < len(row):
                                    try:
                                        val = row[col].get() if hasattr(row[col], 'get') else ''
                                        maxw = max(maxw, cell_font.measure(str(val)) + 12)
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                        # fallback to weight-based if nothing measured
                        if maxw <= 0:
                            maxw = max(18, int((self.col_weights[col]/total) * avail_width))
                        measured.append(maxw)
                    # Normalize measured widths to fit avail_width
                    meas_sum = sum(measured)
                    if meas_sum <= 0:
                        sizes = [max(18, int((w/total)*avail_width)) for w in self.col_weights]
                    else:
                        # Prefer to keep 'Trận' and 'BÀN' narrow: cap their measured widths
                        try:
                            if len(measured) >= 2:
                                # set sensible caps (pixels) for small columns
                                measured[0] = min(measured[0], 80)
                                measured[1] = min(measured[1], 60)
                        except Exception:
                            pass
                        sizes = [max(18, int((m / meas_sum) * avail_width)) for m in measured]
                except Exception:
                    sizes = []
                    for w in self.col_weights:
                        try:
                            minpix = int((w / total) * avail_width)
                            if minpix < 18:
                                minpix = 18
                            sizes.append(minpix)
                        except Exception:
                            sizes.append(18)
                    # Adjust last column to ensure total equals canvas width (avoid small gaps)
                try:
                    ssum = sum(sizes) + total_padding
                    diff = int(canvas_width) - int(ssum)
                    # If some columns are excessively large compared to available area, clamp them
                    try:
                        max_allowed = int(avail_width * 0.40)
                    except Exception:
                        max_allowed = int(canvas_width * 0.6)
                    # Clamp oversized columns (e.g., name columns) and redistribute excess
                    excess = 0
                    for idx, val in enumerate(sizes):
                        if val > max_allowed:
                            excess += (val - max_allowed)
                            sizes[idx] = max_allowed
                    if excess > 0 and len(sizes) > 0:
                        # distribute excess to other columns proportionally (simple loop)
                        per_col_add = excess // len(sizes)
                        for idx in range(len(sizes)):
                            if sizes[idx] < max_allowed:
                                sizes[idx] += per_col_add
                    # final adjustment to absorb any remaining diff
                    ssum = sum(sizes) + total_padding
                    diff = int(canvas_width) - int(ssum)
                    if sizes and diff != 0:
                        # Assign any leftover horizontal pixels to the 'BÀN' column (index 1)
                        try:
                            if len(sizes) > 1:
                                sizes[1] = max(18, sizes[1] + diff)
                            else:
                                sizes[-1] = max(18, sizes[-1] + diff)
                        except Exception:
                            sizes[-1] = max(18, sizes[-1] + diff)
                    # Ensure 'BÀN' occupies at least a fraction of available width so it's visibly larger
                    try:
                        if len(sizes) > 1:
                            # reduce minimum 'BÀN' fraction to 6% so it doesn't dominate layout
                            min_ban = max(18, int(avail_width * 0.06))
                            if sizes[1] < min_ban:
                                sizes[1] = min_ban
                    except Exception:
                        pass
                    # Prefer smaller 'Trận' and compact 'BÀN' columns; apply user-specified reductions
                    try:
                        if len(sizes) >= 2:
                            # Keep Trận small
                            sizes[0] = 40
                            # Apply requested reductions:
                            # - 'BÀN' reduce to 80%
                            # - 'Điểm số' (index 4) reduce to 70%
                            # - 'Địa chỉ vMix' (index 5) reduce to 70%
                            sizes[1] = max(18, int(sizes[1] * 0.8))
                            if len(sizes) > 4:
                                sizes[4] = max(18, int(sizes[4] * 0.7))
                            if len(sizes) > 5:
                                sizes[5] = max(18, int(sizes[5] * 0.7))
                            if len(sizes) >= 4:
                                # Keep name columns reasonable but allow them to take remaining space
                                sizes[2] = max(90, int(sizes[2]))
                                sizes[3] = max(90, int(sizes[3]))
                    except Exception:
                        pass
                    # Enforce explicit pixel caps to ensure rightmost 'Sửa' column remains visible
                    try:
                        # Desired explicit sizes (pixels)
                        desired_ketqua = 50
                        desired_diem = 60
                        desired_vmix = 120
                        # Only apply if sizes length is sufficient
                        if len(sizes) > 6:
                            # compute current sum and replace these three with desired values
                            cur_sum = sum(sizes)
                            cur_three = sizes[4] + sizes[5] + sizes[6]
                            sizes[4] = max(18, desired_diem)
                            sizes[5] = max(18, desired_vmix)
                            sizes[6] = max(18, desired_ketqua)
                            new_sum = sum(sizes)
                            # If we've exceeded available width, reduce name columns (indexes 2 and 3) proportionally
                            total_over = new_sum + total_padding - int(canvas_width)
                            if total_over > 0:
                                # try reducing Tên A/B but keep min 90
                                reducible = 0
                                for idx in (2,3):
                                    if sizes[idx] > 90:
                                        reducible += sizes[idx] - 90
                                if reducible > 0:
                                    need = total_over
                                    for idx in (2,3):
                                        if sizes[idx] > 90:
                                            take = int((sizes[idx]-90)/reducible * need)
                                            sizes[idx] = max(90, sizes[idx] - take)
                                # if still over, shrink 'BÀN' then 'Trận'
                                new_sum2 = sum(sizes) + total_padding
                                if new_sum2 > int(canvas_width):
                                    remain = new_sum2 - int(canvas_width)
                                    if sizes[1] - remain > 18:
                                        sizes[1] = max(18, sizes[1] - remain)
                                    elif sizes[0] - remain > 18:
                                        sizes[0] = max(18, sizes[0] - remain)
                    except Exception:
                        pass
                except Exception:
                    pass
                for i, minpix in enumerate(sizes):
                    try:
                        # Set minsize and avoid allowing grid weights to further expand columns
                        self.table_frame.grid_columnconfigure(i, minsize=minpix, weight=0)
                    except Exception:
                        pass
                # Ensure 'BÀN' column explicitly has the computed minsize (enforce again)
                try:
                    if len(sizes) > 1:
                        self.table_frame.grid_columnconfigure(1, minsize=sizes[1], weight=0)
                except Exception:
                    pass
                # debug logging to help diagnose remaining horizontal gaps
                try:
                    import datetime
                    dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                    with open(dbg, 'a', encoding='utf-8') as df:
                        df.write(f"[{datetime.datetime.now().isoformat()}] _on_canvas_configure canvas_width={canvas_width} avail_width={avail_width} total_padding={total_padding} sizes_sum={sum(sizes)} sizes={sizes}\n")
                except Exception:
                    pass
        except Exception:
            pass

    def clear_table(self):
        for widgets in self.match_rows:
            for w in widgets:
                w.destroy()
        self.match_rows.clear()

    def _normalize_name_for_compare(self, value):
        import re
        import unicodedata
        s = '' if value is None else str(value)
        s = unicodedata.normalize('NFD', s)
        s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
        s = s.lower().strip()
        s = re.sub(r'\s+', ' ', s)
        return s

    def _set_row_position(self, row_idx, swapped):
        try:
            if hasattr(self, '_row_swap_states') and row_idx < len(self._row_swap_states):
                self._row_swap_states[row_idx] = bool(swapped)
        except Exception:
            pass
        try:
            if hasattr(self, '_row_position_vars') and row_idx < len(self._row_position_vars):
                self._row_position_vars[row_idx].set('Ngược' if swapped else 'Thuận')
        except Exception:
            pass
        try:
            if hasattr(self, '_row_swap_buttons') and row_idx < len(self._row_swap_buttons):
                btn = self._row_swap_buttons[row_idx]
                if btn is not None:
                    if swapped:
                        btn.config(bg='#FF5722', text='Đổi ⇄')
                    else:
                        btn.config(bg='#9E9E9E', text='Đổi')
        except Exception:
            pass

    def _toggle_row_swap(self, row_idx):
        if row_idx < 0 or row_idx >= len(self.match_rows):
            return
        widgets = self.match_rows[row_idx]
        if len(widgets) < 4:
            return
        ea = widgets[2]
        eb = widgets[3]
        curr_a = ea.get()
        curr_b = eb.get()
        ea.config(state='normal', fg='#222831')
        ea.delete(0, 'end')
        ea.insert(0, curr_b)
        ea.config(state='readonly', fg='#222831')
        eb.config(state='normal', fg='#222831')
        eb.delete(0, 'end')
        eb.insert(0, curr_a)
        eb.config(state='readonly', fg='#222831')
        current = False
        try:
            if hasattr(self, '_row_swap_states') and row_idx < len(self._row_swap_states):
                current = bool(self._row_swap_states[row_idx])
        except Exception:
            pass
        self._set_row_position(row_idx, not current)

    def _fetch_vmix_livescore_data(self, vmix_url):
        import requests
        import xml.etree.ElementTree as ET
        resp = requests.get(f'{vmix_url}/API/', timeout=3)
        resp.raise_for_status()
        root = ET.fromstring(resp.text)
        input1 = root.find(".//input[@number='1']")

        def get_field(name):
            if input1 is not None:
                for text in input1.findall('text'):
                    if text.attrib.get('name') == name:
                        return (text.text or '').strip()
            return ''

        return {
            'ten_a': get_field('TenA.Text'),
            'ten_b': get_field('TenB.Text'),
            'diem_a': get_field('DiemA.Text'),
            'diem_b': get_field('DiemB.Text'),
            'lco': get_field('Lco.Text'),
            'hr1a': get_field('HR1A.Text'),
            'hr2a': get_field('HR2A.Text'),
            'hr1b': get_field('HR1B.Text'),
            'hr2b': get_field('HR2B.Text'),
        }

    def _post_row_livescore(self, row_idx, vmix_data, silent=False):
        import datetime
        import requests
        import re as _re

        row = self.match_rows[row_idx]
        tran_val = row[0].get().strip()
        event_id = self.event_id_var.get().strip() if hasattr(self, 'event_id_var') else ''
        round_type = self.round_type_var.get() if hasattr(self, 'round_type_var') else 'Vòng Loại'
        base_url = self.hbsf_url_var.get().strip().rstrip('/') if hasattr(self, 'hbsf_url_var') else ''
        if not tran_val or not event_id or not event_id.isdigit() or not base_url:
            return False
        m = _re.search(r'(\d+)', tran_val)
        if not m:
            return False
        match_idx_actual = int(m.group(1))
        if round_type == 'Vòng Loại':
            api_url = f'{base_url}/api/tournament-matches/livescore-update/{event_id}'
            payload = {'match_idx_actual': match_idx_actual}
        else:
            api_url = f'{base_url}/api/main-matches/livescore-update/{event_id}'
            payload = {'match_idx': match_idx_actual}

        swapped = bool(getattr(self, '_row_swap_states', [False] * len(self.match_rows))[row_idx])
        if swapped:
            payload.update({
                'point_A': vmix_data['diem_b'],
                'point_B': vmix_data['diem_a'],
                'turn': vmix_data['lco'],
                'hr1a': vmix_data['hr1b'],
                'hr2a': vmix_data['hr2b'],
                'hr1b': vmix_data['hr1a'],
                'hr2b': vmix_data['hr2a'],
            })
        else:
            payload.update({
                'point_A': vmix_data['diem_a'],
                'point_B': vmix_data['diem_b'],
                'turn': vmix_data['lco'],
                'hr1a': vmix_data['hr1a'],
                'hr2a': vmix_data['hr2a'],
                'hr1b': vmix_data['hr1b'],
                'hr2b': vmix_data['hr2b'],
            })

        resp = requests.post(api_url, json=payload, timeout=5)
        try:
            dbg_path = self._log_path()
            with open(dbg_path, 'a', encoding='utf-8') as df:
                df.write(
                    f"[{datetime.datetime.now()}] KETQUA_API row={row_idx} "
                    f"tran={match_idx_actual} round={round_type} "
                    f"url={api_url} status={resp.status_code} payload={payload} resp={resp.text[:200]}\n"
                )
        except Exception:
            pass
        if resp.status_code != 200:
            return False
        if not silent:
            self.status_var.set(f'Kết quả trận {match_idx_actual} đã đẩy lên web ({round_type})')
        return True

    def _run_ketqua_logic_for_row(self, row_idx, silent=False, show_name_mismatch_popup=True):
        from tkinter import messagebox
        if row_idx < 0 or row_idx >= len(self.match_rows):
            return False
        row = self.match_rows[row_idx]
        tran_val = row[0].get().strip() if len(row) > 0 else ''
        vmix_url = row[5].get().strip() if len(row) > 5 else ''
        if not tran_val or not vmix_url:
            return False
        try:
            vmix_data = self._fetch_vmix_livescore_data(vmix_url)
        except Exception as ex:
            if not silent:
                messagebox.showerror('Lỗi lấy kết quả vMix', f'Không lấy được kết quả từ vMix ({vmix_url}):\n{ex}')
            return False

        screen_a = row[2].get().strip() if len(row) > 2 else ''
        screen_b = row[3].get().strip() if len(row) > 3 else ''
        vmix_a = vmix_data.get('ten_a', '').strip()
        vmix_b = vmix_data.get('ten_b', '').strip()

        n_screen_a = self._normalize_name_for_compare(screen_a)
        n_screen_b = self._normalize_name_for_compare(screen_b)
        n_vmix_a = self._normalize_name_for_compare(vmix_a)
        n_vmix_b = self._normalize_name_for_compare(vmix_b)

        if n_vmix_a and n_vmix_b and n_screen_a and n_screen_b:
            if n_vmix_a == n_screen_b and n_vmix_b == n_screen_a:
                self._toggle_row_swap(row_idx)
            elif not (n_vmix_a == n_screen_a and n_vmix_b == n_screen_b):
                if show_name_mismatch_popup and not silent:
                    messagebox.showwarning('Tên vận động viên không khớp', 'Tên vận động viên không khớp')
                return False
        else:
            if show_name_mismatch_popup and not silent:
                messagebox.showwarning('Tên vận động viên không khớp', 'Tên vận động viên không khớp')
            return False

        try:
            return self._post_row_livescore(row_idx, vmix_data, silent=silent)
        except Exception as ex:
            if not silent:
                messagebox.showerror('Lỗi đẩy kết quả lên web HBSF', str(ex))
            return False

    def _stop_send_blink(self, row_idx):
        try:
            self._send_blink_rows.discard(row_idx)
        except Exception:
            pass
        if row_idx < len(self.match_rows):
            try:
                btn = self.match_rows[row_idx][7]
                btn.config(bg='#00C853', fg='white', text='Gửi')
            except Exception:
                pass
        if not getattr(self, '_send_blink_rows', set()):
            try:
                if getattr(self, '_send_blink_job', None):
                    self.after_cancel(self._send_blink_job)
            except Exception:
                pass
            self._send_blink_job = None

    def _mark_send_needs_refresh(self, row_idx):
        try:
            self._send_blink_rows.add(row_idx)
        except Exception:
            return
        if not getattr(self, '_send_blink_job', None):
            self._blink_send_buttons_tick()

    def _blink_send_buttons_tick(self):
        if not hasattr(self, '_send_blink_rows'):
            return
        self._send_blink_phase = not getattr(self, '_send_blink_phase', False)
        bad_rows = []
        for idx in list(self._send_blink_rows):
            if idx >= len(self.match_rows):
                bad_rows.append(idx)
                continue
            try:
                btn = self.match_rows[idx][7]
                if self._send_blink_phase:
                    btn.config(bg='#D50000', fg='white', text='Gửi !')
                else:
                    btn.config(bg='#8E0000', fg='white', text='Gửi !')
            except Exception:
                bad_rows.append(idx)
        for idx in bad_rows:
            self._send_blink_rows.discard(idx)
        if self._send_blink_rows:
            self._send_blink_job = self.after(450, self._blink_send_buttons_tick)
        else:
            self._send_blink_job = None

    # ── Set Bàn blink methods ─────────────────────────────────────────────────
    def _start_set_ban_blink(self, row_idx):
        try:
            self._set_ban_blink_rows.add(row_idx)
        except Exception:
            return
        if not getattr(self, '_set_ban_blink_job', None):
            self._blink_set_ban_tick()

    def _stop_set_ban_blink(self, row_idx):
        try:
            self._set_ban_blink_rows.discard(row_idx)
        except Exception:
            pass
        try:
            btns = getattr(self, '_set_ban_buttons', [])
            if row_idx < len(btns) and btns[row_idx] is not None:
                btns[row_idx].config(bg='#6A1B9A', fg='white', text='Set Bàn')
        except Exception:
            pass
        if not getattr(self, '_set_ban_blink_rows', set()):
            try:
                if getattr(self, '_set_ban_blink_job', None):
                    self.after_cancel(self._set_ban_blink_job)
            except Exception:
                pass
            self._set_ban_blink_job = None

    def _blink_set_ban_tick(self):
        if not hasattr(self, '_set_ban_blink_rows'):
            return
        self._set_ban_blink_phase = not getattr(self, '_set_ban_blink_phase', False)
        bad_rows = []
        btns = getattr(self, '_set_ban_buttons', [])
        for idx in list(self._set_ban_blink_rows):
            if idx >= len(getattr(self, 'match_rows', [])):
                bad_rows.append(idx)
                continue
            try:
                btn = btns[idx] if idx < len(btns) else None
                if btn is None:
                    bad_rows.append(idx)
                    continue
                if self._set_ban_blink_phase:
                    btn.config(bg='#D50000', fg='white', text='Set Bàn!')
                else:
                    btn.config(bg='#8E0000', fg='white', text='Set Bàn!')
            except Exception:
                bad_rows.append(idx)
        for idx in bad_rows:
            self._set_ban_blink_rows.discard(idx)
        if self._set_ban_blink_rows:
            self._set_ban_blink_job = self.after(450, self._blink_set_ban_tick)
        else:
            self._set_ban_blink_job = None

    def _auto_ketqua_tick(self):
        try:
            if not getattr(self, '_auto_ketqua_enabled', True):
                return
            if getattr(self, '_auto_ketqua_running', False):
                return
            self._auto_ketqua_running = True
            for idx in range(len(getattr(self, 'match_rows', []))):
                self._run_ketqua_logic_for_row(idx, silent=True, show_name_mismatch_popup=False)
        except Exception:
            pass
        finally:
            self._auto_ketqua_running = False
            try:
                self.after(int(getattr(self, 'auto_ketqua_interval_ms', 5000)), self._auto_ketqua_tick)
            except Exception:
                pass

    def populate_table(self):
        # Lưu lại toàn bộ giá trị các ô cũ nếu có
        old_rows = []
        if hasattr(self, 'match_rows') and self.match_rows and not getattr(self, '_restoring_state', False):
            for row in self.match_rows:
                vals = []
                for idx in range(6):
                    if idx < len(row):
                        vals.append(row[idx].get() if hasattr(row[idx], 'get') else '')
                    else:
                        vals.append('')
                old_rows.append(vals)
        self.clear_table()
        # Nếu đang restore thì chỉ tạo widget, không điền giá trị mặc định
        # ...phần lấy dữ liệu Google Sheets và tạo bảng giữ nguyên...
        # Sau khi tạo widgets từng dòng:
        row_bg1 = '#222831'
        row_bg2 = '#393E46'
        num_rows = self.ban_var.get()
        self.match_rows = []
        self.row_states = []
        self._row_swap_states = [False] * num_rows
        self._row_position_vars = []
        self._row_swap_buttons = [None] * num_rows
        self._set_ban_buttons = [None] * num_rows
        try:
            self._send_blink_rows = set()
        except Exception:
            pass
        try:
            self._set_ban_blink_rows = set()
            self._set_ban_blink_job = None
        except Exception:
            pass
        for i in range(num_rows):
            widgets = []
            bg = row_bg1 if i % 2 == 0 else row_bg2
            row_state = {'dirty': False}


            # Hàm chọn màu chữ tương phản với nền (dựa trên độ sáng)
            def contrast_fg(bg_color):
                bg_color = bg_color.lower()
                # Chuyển mã hex sang RGB
                if bg_color.startswith('#'):
                    hex_color = bg_color.lstrip('#')
                    if len(hex_color) == 3:
                        hex_color = ''.join([c*2 for c in hex_color])
                    if len(hex_color) == 6:
                        r = int(hex_color[0:2], 16)
                        g = int(hex_color[2:4], 16)
                        b = int(hex_color[4:6], 16)
                        # Tính độ sáng (perceived luminance)
                        luminance = 0.299*r + 0.587*g + 0.114*b
                        return '#222831' if luminance > 180 else '#fff'
                # Nếu là tên màu
                if bg_color in ['white', 'yellow', 'lightyellow', 'ivory', 'beige']:
                    return '#222831'
                return '#fff'

            # Số trận (giữ nhỏ để tiết kiệm không gian)
            e_tran = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), relief='groove', bd=2, justify='center', insertbackground='white')
            if i < len(old_rows):
                e_tran.insert(0, old_rows[i][0])
            # Reduce internal padding and character width to make the column visually narrower
            e_tran.config(width=4)
            e_tran.grid(row=i+1, column=0, padx=2, pady=2, ipadx=0, ipady=6, sticky='ew')
            try:
                e_tran._prev_tran_value = e_tran.get().strip()
            except Exception:
                pass
            e_tran.bind('<Return>', lambda event, idx=i: self.update_vdv_from_tran(idx))
            e_tran.bind('<FocusOut>', lambda event, idx=i: self.update_vdv_from_tran(idx))
            # Khi bấm Tab, chuyển xuống ô Trận dòng tiếp theo
            def on_tab(event, idx=i):
                if idx+1 < len(self.match_rows):
                    next_tran = self.match_rows[idx+1][0]
                    next_tran.focus_set()
                    next_tran.icursor('end')
                return 'break'  # Chặn chuyển ngang
            e_tran.bind('<Tab>', lambda event, idx=i: on_tab(event, idx))
            # Hiệu ứng dấu nháy trắng khi focus
            def on_focus_in_tran(event, entry=e_tran):
                try:
                    entry.config(insertbackground='white')
                except:
                    pass
            e_tran.bind('<FocusIn>', on_focus_in_tran)
            widgets.append(e_tran)
            # Số bàn — hiển thị dưới dạng nút bấm (nhấn mở popup lịch bàn)
            _ban_var = tk.StringVar(value='')
            if i < len(old_rows):
                _ban_var.set(old_rows[i][1])
            btn_ban = tk.Button(
                self.table_frame,
                textvariable=_ban_var,
                font=('Arial', 14, 'bold'),
                bg='#1565C0', fg='white',
                relief='raised', bd=2,
                width=8,
            )
            # Entry-compatible API (các hàm khác gọi .get()/.delete()/.insert() trên widget này)
            btn_ban.get = lambda v=_ban_var: v.get()
            btn_ban.delete = lambda s, e, v=_ban_var: v.set('')
            btn_ban.insert = lambda pos, val, v=_ban_var: v.set(val)
            btn_ban.config_text = lambda val, v=_ban_var: v.set(val)
            btn_ban.grid(row=i+1, column=1, padx=2, pady=2, ipadx=4, ipady=6, sticky='ew')
            btn_ban.config(command=lambda idx=i, b=btn_ban: self.show_table_schedule_popup(b.get(), idx))
            widgets.append(btn_ban)
            # Tên VĐV A
            e_a = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg='#222831', relief='groove', bd=2)
            if i < len(old_rows):
                e_a.insert(0, old_rows[i][2])
            # Make Tên A column wider
            e_a.config(state='readonly', fg='#222831')
            try:
                # Reduce Tên A visual width to 80% of previous
                e_a.config(width=22)
            except Exception:
                pass
            e_a.grid(row=i+1, column=2, padx=2, pady=2, ipadx=10, ipady=6, sticky='ew')
            widgets.append(e_a)
            # Tên VĐV B
            e_b = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg='#222831', relief='groove', bd=2)
            if i < len(old_rows):
                e_b.insert(0, old_rows[i][3])
            # Make Tên B column wider
            e_b.config(state='readonly', fg='#222831')
            try:
                # Reduce Tên B visual width to 80% of previous
                e_b.config(width=22)
            except Exception:
                pass
            e_b.grid(row=i+1, column=3, padx=2, pady=2, ipadx=10, ipady=6, sticky='ew')
            widgets.append(e_b)
            # Điểm số (giảm width)
            e_diem = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), justify='center', relief='groove', bd=2)
            if i < len(old_rows):
                e_diem.insert(0, old_rows[i][4])
            try:
                e_diem.config()
            except Exception:
                pass
            e_diem.config(width=3)
            e_diem.grid(row=i+1, column=4, padx=2, pady=2, ipadx=0, ipady=6, sticky='ew')
            def diem_user_edit(event, entry=e_diem):
                entry._user_edited = True
            e_diem.bind('<Key>', diem_user_edit)
            widgets.append(e_diem)
            # Địa chỉ vMix cho từng dòng (giữ lại giá trị cũ nếu có)
            e_vmix = tk.Entry(self.table_frame, font=('Arial', 16), bg='#B3E5FC', fg='#222831', relief='groove', bd=2)
            if i < len(old_rows) and old_rows[i][5]:
                e_vmix.insert(0, old_rows[i][5])
            else:
                e_vmix.insert(0, 'http://127.0.0.1:8088')
            try:
                e_vmix.config()
            except Exception:
                pass
            e_vmix.grid(row=i+1, column=5, padx=2, pady=2, ipadx=0, ipady=6, sticky='ew')
            widgets.append(e_vmix)
            # Nút Kết quả cập nhật web cho từng bàn
            swap_state = [False]  # Trạng thái đảo vị trí VĐV: False = bình thường, True = A↔B
            self._set_row_position(i, swap_state[0])
            # --- Nút Kết quả ---
            btn_ketqua = tk.Button(self.table_frame, text='Kết quả', bg='#00C853', fg='white', font=('Arial', 18, 'bold'),
                                   relief='raised', bd=2, width=10)
            btn_ketqua.config(width=6)
            btn_ketqua.original_text = 'Kết quả'
            btn_ketqua.grid(row=i+1, column=6, padx=2, pady=2, ipadx=0)
            widgets.append(btn_ketqua)
            # --- Nút Gửi ---
            btn_gui = tk.Button(self.table_frame, text='Gửi', bg='#00C853', fg='white', font=('Arial', 18, 'bold'),
                               relief='raised', bd=2, width=10)
            btn_gui.config(width=6)
            btn_gui.original_text = 'Gửi'
            btn_gui.grid(row=i+1, column=7, padx=2, pady=2, ipadx=0)
            widgets.append(btn_gui)
            # --- Nút Sửa ---
            btn_sua = tk.Button(self.table_frame, text='Sửa', bg='#FF9800', fg='#222831', font=('Arial', 18, 'bold'),
                               command=lambda idx=i: self.open_edit_popup(idx), relief='raised', bd=2, width=10)
            btn_sua.config(width=6, disabledforeground='#222831')
            btn_sua.original_text = 'Sửa'
            btn_sua.grid(row=i+1, column=8, padx=2, pady=2, ipadx=0)
            widgets.append(btn_sua)

            # --- Logic đổi màu nút ---
            def set_btn_color(btn, state):
                # Chỉ nút Sửa mới đổi text, các nút khác giữ nguyên text gốc
                if hasattr(btn, 'original_text') and btn.original_text == 'Sửa':
                    if state == 'pending':
                        btn.config(bg='#00C853', fg='white', text='Sửa', state='normal')  # Xanh lá
                    elif state == 'edit':
                        btn.config(bg='#388E3C', fg='white', text='Sửa', state='normal')  # Xanh đậm khi sửa
                    elif state == 'success':
                        btn.config(bg='#D32F2F', fg='white', text='Sửa', state='normal')  # Đỏ khi gửi thành công
                    elif state == 'fail':
                        btn.config(bg='black', fg='white', text='Sửa', state='normal')    # Đen khi lỗi nhưng vẫn cho sửa lại
                else:
                    # Các nút khác chỉ đổi màu, giữ nguyên text
                    if state == 'pending':
                        btn.config(bg='#00C853', fg='white', state='normal')
                    elif state == 'edit':
                        btn.config(bg='#388E3C', fg='white', state='normal')
                    elif state == 'success':
                        btn.config(bg='#D32F2F', fg='white', state='normal')
                    elif state == 'fail':
                        btn.config(bg='black', fg='white', state='normal')

            set_btn_color(btn_gui, 'pending')
            set_btn_color(btn_ketqua, 'pending')
            btn_sua.config(text='Sửa', state='normal', bg='#FF9800', fg='#222831')

            # --- Theo dõi thay đổi tên A/B để đổi màu nút ---
            def on_edit_name(event=None, btns=(btn_gui, btn_ketqua)):
                for b in btns:
                    set_btn_color(b, 'edit')
            widgets[2].bind('<KeyRelease>', on_edit_name)
            widgets[3].bind('<KeyRelease>', on_edit_name)

            # --- Nút Gửi: gửi xong đổi màu ---
            def on_push_to_vmix(idx=i, btn=btn_gui):
                self._stop_send_blink(idx)
                try:
                    self.push_to_vmix(idx)
                    set_btn_color(btn, 'pending')
                except Exception:
                    set_btn_color(btn, 'fail')
            btn_gui.config(command=lambda idx=i, btn=btn_gui: on_push_to_vmix(idx, btn))

            # --- Nút Kết quả: gửi xong đổi màu ---
            def on_btn_ketqua(idx=i, btn=btn_ketqua):
                try:
                    ok = self._run_ketqua_logic_for_row(idx, silent=False, show_name_mismatch_popup=True)
                    set_btn_color(btn, 'success' if ok else 'fail')
                except Exception:
                    set_btn_color(btn, 'fail')
            btn_ketqua.config(command=lambda idx=i, btn=btn_ketqua: on_btn_ketqua(idx, btn))

            # (Removed per-row 'Đổi', 'Vị Trí' and 'Set Bàn' controls — layout adjusted to avoid empty columns)
            self._set_row_position(i, swap_state[0])
            # Không còn bôi màu khi click vào ô
            self.match_rows.append(widgets)
            self.row_states.append(row_state)
        self.selected_row = None

    def update_vdv_from_tran(self, row_idx):
        """Khi nhập số trận, tự động lấy tên VĐV từ sheet và cập nhật dòng giao diện."""
        import sys
        # Chỉ dùng self.sheet_rows đã được cập nhật mới nhất
        current_rows = getattr(self, 'sheet_rows', [])
        if not current_rows or not isinstance(current_rows, list) or not current_rows[0]:
            self.status_var.set('Chưa tải dữ liệu trận đấu! (Dùng nút "Lấy thông tin trận đấu")')
            print('DEBUG: sheet_rows is empty or invalid', file=sys.stderr)
            return
        widgets = self.match_rows[row_idx]
        tran_val = widgets[0].get().strip()
        try:
            prev_tran = getattr(widgets[0], '_prev_tran_value', tran_val)
            if tran_val != prev_tran:
                if tran_val:
                    self._mark_send_needs_refresh(row_idx)
                else:
                    self._stop_send_blink(row_idx)
                widgets[0]._prev_tran_value = tran_val
        except Exception:
            pass
        # Kiểm tra trùng số trận (số bàn không kiểm tra vì nhiều trận có thể dùng chung bàn)
        is_tran_duplicate = False
        for i, row in enumerate(self.match_rows):
            if i != row_idx:
                if row[0].get().strip() == tran_val and tran_val:
                    is_tran_duplicate = True
        def set_status(msg):
            try:
                self.status_var.set(msg)
            except Exception:
                pass
        if is_tran_duplicate:
            from tkinter import messagebox
            widgets[0].delete(0, 'end')
            try:
                widgets[0]._prev_tran_value = ''
            except Exception:
                pass
            self._stop_send_blink(row_idx)
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            set_status('Số trận đã bị trùng!')
            messagebox.showwarning('Cảnh báo', 'Số trận này đã bị trùng! Vui lòng nhập số khác.')
            widgets[0].focus_set()
            return
        # Nếu số trận bị trùng thì không lấy tên VĐV, xoá ô tên VĐV
        if hasattr(widgets[0], 'is_duplicate') and widgets[0].is_duplicate:
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            return
        # Khi cập nhật số trận mới, trả lại màu xanh lá cho nút Gửi và Kết quả
        if len(widgets) > 7:
            btn_ketqua = widgets[6]
            btn_gui = widgets[7]
            def set_btn_color(btn, state):
                if hasattr(btn, 'original_text') and btn.original_text == 'Sửa':
                    if state == 'pending':
                        btn.config(bg='#00C853', fg='white', text='Sửa', state='normal')
                    elif state == 'edit':
                        btn.config(bg='#388E3C', fg='white', text='Sửa', state='normal')
                    elif state == 'success':
                        btn.config(bg='#D32F2F', fg='white', text='Sửa', state='normal')
                    elif state == 'fail':
                        btn.config(bg='black', fg='white', text='Sửa', state='normal')
                else:
                    if state == 'pending':
                        btn.config(bg='#00C853', fg='white', state='normal')
                    elif state == 'edit':
                        btn.config(bg='#388E3C', fg='white', state='normal')
                    elif state == 'success':
                        btn.config(bg='#D32F2F', fg='white', state='normal')
                    elif state == 'fail':
                        btn.config(bg='black', fg='white', state='normal')
            set_btn_color(btn_gui, 'pending')
            set_btn_color(btn_ketqua, 'pending')
        # Nếu ô Trận bị xóa trắng hoặc bị xóa ký tự (delete/backspace) thì xóa luôn tên VĐV
        if not tran_val:
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            # Luôn giữ nút Sửa ổn định, không đổi text
            if len(widgets) > 8:
                widgets[8].config(state='normal', text='Sửa', bg='#FF9800', fg='#222831')
            try:
                self._stop_set_ban_blink(row_idx)
            except Exception:
                pass
            return
        # Tìm tên cột linh hoạt, cho phép chọn cột theo tên thực tế
        def find_col_key(keys, *candidates):
            def normalize(s):
                return s.replace(' ', '').replace('_', '').lower()
            norm_keys = {normalize(k): k for k in keys}
            for c in candidates:
                nc = normalize(c)
                for nk, orig in norm_keys.items():
                    if nk == nc:
                        return orig
            # Nếu không tìm thấy, thử lấy theo vị trí cột (B, C, D, F, H...)
            col_letters = [c for c in candidates if len(c) == 1 and c.isalpha()]
            if col_letters:
                idx = ord(col_letters[0].upper()) - ord('A')
                if idx < len(keys):
                    return keys[idx]
            return None
        # Chỉ dùng self.sheet_rows, không gọi lại fetch_matches_from_ketqua
        source_rows = current_rows if current_rows else []
        if not source_rows:
            set_status('Không có dữ liệu Google Sheet!')
            print('DEBUG: Không có dữ liệu Google Sheet!', file=sys.stderr)
            return
        keys = list(source_rows[0].keys())
        print(f'DEBUG: source_rows[0] keys = {keys}', file=sys.stderr)
        # Cho phép người dùng chỉnh tên cột tại đây (tìm trong sheet 'KET QUA'):
        tran_col = find_col_key(keys, 'Trận', 'Số trận', 'B', 'C')
        vdv_a_col = find_col_key(keys, 'VĐVA', 'vđv a', 'D', 'F')
        vdv_b_col = find_col_key(keys, 'VĐVB', 'vđv b', 'F', 'H')
        print(f'DEBUG: tran_col={tran_col}, vdv_a_col={vdv_a_col}, vdv_b_col={vdv_b_col}', file=sys.stderr)
        if not tran_col or not vdv_a_col or not vdv_b_col:
            set_status('Không tìm thấy tiêu đề cột!')
            print('DEBUG: Không tìm thấy tiêu đề cột!', file=sys.stderr)
            widgets[2].config(state='normal'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly')
            widgets[3].config(state='normal'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly')
            return
        # So sánh số trận dạng số nguyên nếu có thể
        import re
        def normalize_tran(val):
            import re
            s = str(val)
            # Lấy tất cả số liên tiếp đầu tiên trong chuỗi
            m = re.search(r'(\d+)', s)
            if m:
                return int(m.group(1))
            return s.strip().lower()
        input_tran_num = normalize_tran(tran_val)
        print(f'DEBUG: input_tran_num={input_tran_num}', file=sys.stderr)
        found = None
        for idx, r in enumerate(source_rows):
            sheet_tran_val = r.get(tran_col, '')
            sheet_tran_num = normalize_tran(sheet_tran_val)
            print(f'DEBUG: row {idx} {tran_col}={sheet_tran_val!r} -> {sheet_tran_num!r} | input={input_tran_num!r}', file=sys.stderr)
            if sheet_tran_num == input_tran_num:
                found = r
                print(f'DEBUG: Found match at row {idx}', file=sys.stderr)
                break
        if not found:
            print(f'DEBUG: Không tìm thấy trận. tran_val nhập vào: {tran_val!r}, input_tran_num: {input_tran_num!r}', file=sys.stderr)
            print(f'DEBUG: tran_col={tran_col}, sheet_rows keys: {[list(r.keys()) for r in source_rows]}', file=sys.stderr)
        if not found:
            print('DEBUG: Không tìm thấy trận phù hợp', file=sys.stderr)
            # Nếu không tìm thấy thì xóa tên VĐV, báo lỗi rõ ràng
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            set_status('Không tìm thấy trận này trong dữ liệu đã tải!')
            if len(widgets) > 8:
                widgets[8].config(state='normal', text='Sửa', bg='#FF9800', fg='#222831')
        else:
            vdv_a = found.get(vdv_a_col, '') if vdv_a_col else ''
            vdv_b = found.get(vdv_b_col, '') if vdv_b_col else ''
            # Luôn để màu chữ tên VĐV là đen
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].insert(0, vdv_a); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].insert(0, vdv_b); widgets[3].config(state='readonly', fg='#222831')
            # Nếu hàng đang ở trạng thái 'Đổi', hoán đổi lại tên VĐV
            if getattr(self, '_row_swap_states', None) and row_idx < len(self._row_swap_states) and self._row_swap_states[row_idx]:
                _tmp = widgets[2].get()
                widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].insert(0, widgets[3].get()); widgets[2].config(state='readonly', fg='#222831')
                widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].insert(0, _tmp); widgets[3].config(state='readonly', fg='#222831')
            try:
                self._set_row_position(row_idx, bool(getattr(self, '_row_swap_states', [False])[row_idx]))
            except Exception:
                pass
            # Lấy tên bàn mà trận này được xếp vào (từ dữ liệu HBSF)
            ban_val_from_data = found.get('Số bàn', '')
            if not ban_val_from_data:
                ban_col = find_col_key(keys, 'Số bàn', 'Ban', 'SoBan')
                if ban_col:
                    ban_val_from_data = found.get(ban_col, '')
            print(f'DEBUG ban: tran={input_tran_num} ban_val_from_data={ban_val_from_data!r} row_ban={widgets[1].get().strip()!r}', file=sys.stderr)
            if getattr(self, 'hbsf_tables', None):
                # Chế độ HBSF: validate bàn của trận phải khớp với bàn của dòng hiện tại
                row_ban = widgets[1].get().strip()
                if ban_val_from_data and row_ban and str(ban_val_from_data).strip() != str(row_ban).strip():
                    from tkinter import messagebox
                    widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
                    widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
                    widgets[0].delete(0, 'end')
                    set_status(f'Trận #{tran_val} được xếp vào {ban_val_from_data}, không phải {row_ban}!')
                    messagebox.showwarning('Sai bàn', f'Trận #{tran_val} được xếp vào {ban_val_from_data}, không phải {row_ban}!\nVui lòng nhập số trận đúng với bàn này.')
                    widgets[0].focus_set()
                    return
                # Nếu hợp lệ: giữ nguyên tên bàn đã pre-filled, không ghi đè
            else:
                # Chế độ cũ (Google Sheet): điền bàn nếu ô đang trống
                if ban_val_from_data and not widgets[1].get().strip():
                    widgets[1].delete(0, 'end')
                    widgets[1].insert(0, ban_val_from_data)
            set_status('')
            # Nếu trận chưa có bàn (HBSF mode) → nhấp nháy nút Set Bàn
            try:
                if getattr(self, 'hbsf_tables', None):
                    if not ban_val_from_data:
                        self._start_set_ban_blink(row_idx)
                    else:
                        self._stop_set_ban_blink(row_idx)
                else:
                    self._stop_set_ban_blink(row_idx)
            except Exception:
                pass
            if len(widgets) > 8:
                widgets[8].config(state='normal', text='Sửa', bg='#FF9800', fg='#222831')

    def _log_path(self):
        """Trả về đường dẫn tuyệt đối tới file log, ưu tiên cạnh exe/script."""
        import os
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            base = os.getcwd()
        return os.path.join(base, 'vmix_debug.log')

    def show_log_popup(self):
        """Mở popup hiển thị 100 dòng cuối của vmix_debug.log."""
        import tkinter as tk
        import os

        log_path = self._log_path()
        if not os.path.exists(log_path):
            from tkinter import messagebox
            messagebox.showinfo('Log', f'Chưa có file log.\nĐường dẫn: {log_path}')
            return

        try:
            with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
                lines = f.readlines()
        except Exception as ex:
            from tkinter import messagebox
            messagebox.showerror('Lỗi', f'Không đọc được log: {ex}')
            return

        last_lines = lines[-100:]

        top = tk.Toplevel(self)
        top.title(f'Log — {log_path}')
        top.configure(bg='#1a1a2e')
        top.geometry('900x520')
        top.resizable(True, True)

        header_frame = tk.Frame(top, bg='#1a1a2e')
        header_frame.pack(fill='x', padx=10, pady=(8, 2))
        tk.Label(header_frame, text=f'vmix_debug.log  ({len(lines)} dòng, hiển thị {len(last_lines)} dòng cuối)',
                 font=('Consolas', 10), bg='#1a1a2e', fg='#FFD369').pack(side='left')
        tk.Button(header_frame, text='Xoá log', font=('Segoe UI', 10),
                  bg='#b71c1c', fg='white', relief='flat',
                  command=lambda: self._clear_log(log_path, text_widget)).pack(side='right', padx=4)
        tk.Button(header_frame, text='Làm mới', font=('Segoe UI', 10),
                  bg='#1565C0', fg='white', relief='flat',
                  command=lambda: self._refresh_log(log_path, text_widget)).pack(side='right', padx=4)

        text_frame = tk.Frame(top, bg='#1a1a2e')
        text_frame.pack(fill='both', expand=True, padx=10, pady=4)

        vsb = tk.Scrollbar(text_frame)
        vsb.pack(side='right', fill='y')
        hsb = tk.Scrollbar(text_frame, orient='horizontal')
        hsb.pack(side='bottom', fill='x')

        text_widget = tk.Text(
            text_frame, font=('Consolas', 10), bg='#0d1117', fg='#c9d1d9',
            wrap='none', relief='flat',
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        text_widget.pack(fill='both', expand=True)
        vsb.config(command=text_widget.yview)
        hsb.config(command=text_widget.xview)

        # Màu dòng theo nội dung
        text_widget.tag_config('ok',   foreground='#56d364')
        text_widget.tag_config('err',  foreground='#f85149')
        text_widget.tag_config('warn', foreground='#e3b341')

        for line in last_lines:
            if '✅' in line or 'status=200' in line:
                tag = 'ok'
            elif '❌' in line or 'ERROR' in line or 'Lỗi' in line:
                tag = 'err'
            elif 'rowCount=0' in line or 'WARN' in line:
                tag = 'warn'
            else:
                tag = ''
            text_widget.insert('end', line, tag)

        text_widget.see('end')
        text_widget.config(state='disabled')

        tk.Button(top, text='Đóng', font=('Segoe UI', 11), bg='#555', fg='white',
                  command=top.destroy).pack(pady=6)

    def _refresh_log(self, log_path, text_widget):
        """Tải lại nội dung log vào text_widget."""
        import os
        text_widget.config(state='normal')
        text_widget.delete('1.0', 'end')
        if os.path.exists(log_path):
            try:
                with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
                    lines = f.readlines()
                for line in lines[-100:]:
                    if '✅' in line or 'status=200' in line:
                        tag = 'ok'
                    elif '❌' in line or 'Lỗi' in line:
                        tag = 'err'
                    elif 'rowCount=0' in line or 'WARN' in line:
                        tag = 'warn'
                    else:
                        tag = ''
                    text_widget.insert('end', line, tag)
                text_widget.see('end')
            except Exception:
                pass
        text_widget.config(state='disabled')

    def _clear_log(self, log_path, text_widget):
        """Xoá toàn bộ nội dung file log."""
        from tkinter import messagebox
        if not messagebox.askyesno('Xác nhận', 'Xoá toàn bộ log?'):
            return
        try:
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write('')
            self._refresh_log(log_path, text_widget)
        except Exception as ex:
            messagebox.showerror('Lỗi', f'Không xoá được log: {ex}')

    def show_table_schedule_popup(self, table_name, row_idx):
        """Hiển thị popup danh sách trận của một bàn cụ thể."""
        import tkinter as tk
        hbsf_data = getattr(self, 'hbsf_match_data', {})
        if not hbsf_data:
            from tkinter import messagebox
            messagebox.showinfo('Thông tin', 'Chưa tải dữ liệu HBSF. Nhấn "Lấy thông tin HBSF" trước.')
            return

        if table_name:
            table_matches = [
                {'match_idx': k, **v}
                for k, v in hbsf_data.items()
                if v.get('table_name', '') == table_name
            ]
        else:
            table_matches = [
                {'match_idx': k, **v}
                for k, v in hbsf_data.items()
                if not v.get('table_name', '')
            ]

        table_matches.sort(key=lambda m: (not m.get('match_time'), m.get('match_time') or ''))

        top = tk.Toplevel(self)
        top.title(f'Lịch thi đấu – {table_name or "Chưa có bàn"}')
        top.configure(bg='#1a1a2e')
        top.geometry('680x440')
        top.resizable(True, True)

        tk.Label(
            top,
            text=f'Lịch thi đấu: {table_name or "Chưa có bàn"}',
            font=('Arial', 14, 'bold'), bg='#1a1a2e', fg='#FFD369',
        ).pack(pady=(10, 4))

        container = tk.Frame(top, bg='#1a1a2e')
        container.pack(fill='both', expand=True, padx=10, pady=4)

        canvas = tk.Canvas(container, bg='#1a1a2e', bd=0, highlightthickness=0)
        vsb = tk.Scrollbar(container, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        inner = tk.Frame(canvas, bg='#1a1a2e')
        inner_win = canvas.create_window((0, 0), window=inner, anchor='nw')

        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
            canvas.itemconfig(inner_win, width=canvas.winfo_width())
        inner.bind('<Configure>', on_configure)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(inner_win, width=e.width))

        hdrs = ['Trận', 'Giờ', 'VĐV A', 'VĐV B', 'Trạng thái']
        col_w = [6, 16, 20, 20, 12]
        for c, (h, w) in enumerate(zip(hdrs, col_w)):
            tk.Label(
                inner, text=h, font=('Arial', 10, 'bold'),
                bg='#2c2c54', fg='#FFD369',
                width=w, anchor='w', padx=4, pady=4, relief='flat',
            ).grid(row=0, column=c, sticky='ew', padx=1, pady=1)

        for r, m in enumerate(table_matches):
            is_done = m.get('winner_id') is not None
            fg = '#888888' if is_done else '#ffffff'
            row_bg = '#111122' if r % 2 == 0 else '#1a1a2e'
            status = '✓ Hoàn thành' if is_done else '⏳ Chờ thi đấu'
            mtime = m.get('match_time') or ''
            mtime_disp = mtime[11:16] if len(mtime) >= 16 else mtime
            vals = [m.get('match_idx', ''), mtime_disp, m.get('player_a', ''), m.get('player_b', ''), status]
            for c, (v, w) in enumerate(zip(vals, col_w)):
                tk.Label(
                    inner, text=str(v), font=('Arial', 10),
                    bg=row_bg, fg=fg,
                    width=w, anchor='w', padx=4, pady=3, relief='flat',
                ).grid(row=r + 1, column=c, sticky='ew', padx=1, pady=0)

        tk.Button(top, text='Đóng', font=('Arial', 11), bg='#555', fg='white',
                  command=top.destroy).pack(pady=8)

    def update_table_for_row(self, row_idx):
        """Gửi trực tiếp số trận + tên bàn lên web để cập nhật table_id nếu trận chưa có bàn."""
        import re
        from tkinter import messagebox
        try:
            import requests
            widgets = self.match_rows[row_idx]
            tran_val = widgets[0].get().strip()
            table_name = widgets[1].get().strip()
            base_url = self.hbsf_url_var.get().strip().rstrip('/') if hasattr(self, 'hbsf_url_var') else ''
            event_id = self.event_id_var.get().strip() if hasattr(self, 'event_id_var') else ''
            round_type = self.round_type_var.get() if hasattr(self, 'round_type_var') else 'Vòng Loại'

            if not tran_val:
                messagebox.showwarning('Thiếu thông tin', 'Vui lòng nhập số trận trước khi set bàn.')
                return
            if not table_name:
                messagebox.showwarning('Thiếu thông tin', 'Vui lòng nhập/chọn tên bàn trước khi set bàn.')
                return
            if not base_url:
                messagebox.showwarning('Thiếu thông tin', 'Chưa nhập URL web (HBSF URL).')
                return
            if not event_id or not event_id.isdigit():
                messagebox.showwarning('Thiếu thông tin', 'Chưa nhập Event ID hợp lệ.')
                return

            mx = re.search(r'(\d+)', tran_val)
            if not mx:
                messagebox.showerror('Lỗi', f'Không đọc được số trận từ "{tran_val}".')
                return
            match_idx_num = int(mx.group(1))

            if round_type == 'Vòng Loại':
                url = f'{base_url}/api/tournament-matches/table-update/{event_id}'
                payload = {
                    'match_idx_actual': match_idx_num,
                    'table_name': table_name,
                }
            else:
                url = f'{base_url}/api/main-matches/table-update/{event_id}'
                payload = {
                    'match_idx': match_idx_num,
                    'table_name': table_name,
                }

            resp = requests.post(url, json=payload, timeout=5)
            if resp.status_code == 200:
                resp_data = {}
                try:
                    resp_data = resp.json()
                except Exception:
                    pass
                skipped = resp_data.get('skipped', False)
                # Cập nhật hbsf_match_data local
                try:
                    hbsf_data = getattr(self, 'hbsf_match_data', {})
                    key = str(match_idx_num)
                    if key in hbsf_data:
                        hbsf_data[key]['table_name'] = table_name
                except Exception:
                    pass
                # Dừng blink sau khi set thành công
                try:
                    self._stop_set_ban_blink(row_idx)
                except Exception:
                    pass
                if skipped:
                    self.status_var.set(f'Trận {match_idx_num} đã có bàn sẵn, không thay đổi.')
                    messagebox.showinfo(
                        'Đã có bàn',
                        f'Trận {match_idx_num} đã được xếp vào bàn trước đó.\nKhông thay đổi.'
                    )
                else:
                    self.status_var.set(f'✅ Đã set bàn "{table_name}" cho trận {match_idx_num}.')
                    messagebox.showinfo(
                        'Set Bàn thành công',
                        f'✅ Đã cập nhật bàn "{table_name}" cho trận {match_idx_num} ({round_type}).'
                    )
            else:
                try:
                    err_msg = resp.json().get('message', resp.text[:300])
                except Exception:
                    err_msg = resp.text[:300]
                self.status_var.set(f'Set bàn lỗi ({resp.status_code}): {err_msg}')
                messagebox.showerror(
                    'Set Bàn thất bại',
                    f'❌ Không thể set bàn "{table_name}" cho trận {match_idx_num}.\n\nLỗi ({resp.status_code}): {err_msg}'
                )
        except Exception as ex:
            self.status_var.set(f'Lỗi kết nối set bàn: {ex}')
            messagebox.showerror('Lỗi kết nối', f'Không thể kết nối để set bàn:\n{ex}')

    def highlight_row(self, row_idx):
        for i, widgets in enumerate(self.match_rows):
            for j, w in enumerate(widgets):
                if i == row_idx:
                    w.config(bg='#FFD369')
                else:
                    bg = '#222831' if i % 2 == 0 else '#393E46'
                    w.config(bg=bg)
                # Luôn giữ màu chữ ô tên VĐV là đen
                if j == 2 or j == 3:
                    w.config(fg='#222831')
        self.selected_row = row_idx

    def load_hbsf_matches(self):
        """Lấy thông tin trận đấu từ website HBSF qua API và điền vào bảng.

        Dữ liệu được chuyển sang định dạng sheet_rows để tương thích với
        update_vdv_from_tran: mỗi hàng có key 'Trận', 'VĐVA', 'VĐVB', 'Số bàn'.
        """
        import requests
        import sys

        event_id = self.event_id_var.get().strip() if hasattr(self, 'event_id_var') else ''
        round_type = self.round_type_var.get() if hasattr(self, 'round_type_var') else 'Vòng Loại'
        base_url = self.hbsf_url_var.get().strip() if hasattr(self, 'hbsf_url_var') else 'https://hbsf.com.vn'
        base_url = base_url.rstrip('/')

        print(f"[HBSF] load_hbsf_matches called: event_id={event_id!r}, round_type={round_type!r}, base_url={base_url!r}", file=sys.stderr)

        if not event_id:
            self.status_var.set('Vui lòng nhập Event ID!')
            print("[HBSF] Abort: event_id trống", file=sys.stderr)
            return
        if not event_id.isdigit():
            self.status_var.set('Event ID phải là số nguyên!')
            print(f"[HBSF] Abort: event_id không phải số: {event_id!r}", file=sys.stderr)
            return

        api_type = 'qualifier' if round_type == 'Vòng Loại' else 'main'
        url = f"{base_url}/api/tournament-matches/livescore/{event_id}?type={api_type}"

        print(f"[HBSF] Gọi API: {url}", file=sys.stderr)
        self.status_var.set(f'Đang tải trận đấu từ HBSF (Event {event_id}, {round_type})...')
        try:
            resp = requests.get(url, timeout=10)
        except Exception as e:
            msg = f'Không kết nối được HBSF: {e}'
            print(f"[HBSF] Lỗi kết nối: {e}", file=sys.stderr)
            self.status_var.set(msg)
            try:
                from tkinter import messagebox
                messagebox.showerror('HBSF API Error', msg)
            except Exception:
                pass
            return

        print(f"[HBSF] Response: status={resp.status_code}", file=sys.stderr)
        if resp.status_code != 200:
            msg = f'HBSF API trả về lỗi {resp.status_code}: {resp.text[:200]}'
            print(f"[HBSF] Lỗi HTTP: {msg}", file=sys.stderr)
            self.status_var.set(msg)
            try:
                from tkinter import messagebox
                messagebox.showerror('HBSF API Error', msg)
            except Exception:
                pass
            return

        try:
            data = resp.json()
        except Exception as e:
            self.status_var.set(f'Lỗi phân tích JSON từ HBSF: {e}')
            return

        # Lấy num_tables từ response để cập nhật Tổng Số Bàn
        num_tables = data.get('num_tables')
        print(f"[HBSF] num_tables={num_tables}", file=sys.stderr)

        matches = data.get('matches', [])
        print(f"[HBSF] Số trận nhận được: {len(matches)}", file=sys.stderr)

        # Chuyển sang định dạng sheet_rows tương thích với update_vdv_from_tran
        # Key 'Trận' = match_idx_actual, 'VĐVA'/'VĐVB' = tên VĐV, 'Số bàn' = tên bàn
        converted = []
        for m in matches:
            # Ưu tiên table_name từ DB; fallback dùng table_id nếu table_name là null
            tname = m.get('table_name') or ''
            if not tname:
                tid = m.get('table_id')
                if tid is not None and str(tid).strip() not in ('', '0', 'None'):
                    try:
                        tname = f'Bàn {int(float(str(tid)))}'
                    except (ValueError, TypeError):
                        pass
            converted.append({
                'Trận':    str(m.get('match_idx_actual', '')),
                'VĐVA':    m.get('player_name_a', '') or '',
                'VĐVB':    m.get('player_name_b', '') or '',
                'Số bàn':  tname,
            })

        # Lưu dữ liệu đầy đủ của từng trận để phục vụ popup và cập nhật bàn
        self.hbsf_match_data = {}
        for m in matches:
            tname = m.get('table_name') or ''
            if not tname:
                tid = m.get('table_id')
                if tid is not None and str(tid).strip() not in ('', '0', 'None'):
                    try:
                        tname = f'Bàn {int(float(str(tid)))}'
                    except (ValueError, TypeError):
                        pass
            self.hbsf_match_data[str(m.get('match_idx_actual', ''))] = {
                'id': m.get('id'),
                'match_idx_actual': m.get('match_idx_actual'),
                'match_idx': m.get('match_idx'),
                'table_name': tname,
                'player_a': m.get('player_name_a', '') or '',
                'player_b': m.get('player_name_b', '') or '',
                'match_time': m.get('match_time', '') or '',
                'winner_id': m.get('winner_id'),
            }

        self.sheet_rows = converted

        # Lưu danh sách bàn theo thứ tự index từ API (dùng để validate khi nhập số trận)
        tables_data = data.get('tables', [])
        if tables_data:
            # Backend mới: trả về danh sách bàn theo index tăng dần
            self.hbsf_tables = [t.get('name', '') for t in tables_data]
            # Map tên bàn -> id để dùng khi gọi API cập nhật bàn
            self.hbsf_table_map = {t.get('name', ''): t.get('id') for t in tables_data}
        else:
            # Fallback cho backend cũ: trích xuất thứ tự bàn từ danh sách trận
            # Scheduler xếp bàn xoay vòng nên num_tables trận đầu tiên bao đủ thứ tự
            n = int(num_tables) if num_tables else 0
            seen, seen_set = [], set()
            for m in converted:
                tname = m.get('Số bàn', '')
                if tname and tname not in seen_set:
                    seen.append(tname)
                    seen_set.add(tname)
                    if n and len(seen) >= n:
                        break
            self.hbsf_tables = seen
            self.hbsf_table_map = {}  # Fallback: không có ID bàn
        print(f"[HBSF] hbsf_tables={self.hbsf_tables}", file=sys.stderr)

        # Cập nhật Tổng Số Bàn từ server rồi rebuild bảng để đúng số dòng
        if num_tables is not None and hasattr(self, 'ban_var') and self.ban_var is not None:
            try:
                self.ban_var.set(int(num_tables))
                print(f"[HBSF] Đã set ban_var = {num_tables}", file=sys.stderr)
            except (ValueError, TypeError) as e:
                print(f"[HBSF] Lỗi set ban_var: {e}", file=sys.stderr)

        # Rebuild bảng (cập nhật số dòng và khôi phục giá trị cũ)
        self.populate_table()

        # Pre-fill tên bàn cho từng dòng theo thứ tự index từ HBSF
        for i, row_widgets in enumerate(self.match_rows):
            if i < len(self.hbsf_tables) and self.hbsf_tables[i]:
                row_widgets[1].delete(0, 'end')
                row_widgets[1].insert(0, self.hbsf_tables[i])

        # Tự động điền tên VĐV cho tất cả hàng đang có số Trận
        for i in range(len(self.match_rows)):
            if self.match_rows[i][0].get().strip():
                self.update_vdv_from_tran(i)

        count = len(converted)
        status_parts = [f'Đã tải {count} trận từ HBSF ({round_type}, Event {event_id})']
        if num_tables is not None:
            status_parts.append(f'Tổng số bàn: {num_tables}')
        self.status_var.set('. '.join(status_parts) + '.')

    def reload_matches(self):
        current_url = self.url_var.get().strip() if hasattr(self, 'url_var') else ''
        try:
            print(f"DEBUG: Reloading matches from sheet_url={current_url}", file=sys.stderr)
        except Exception:
            pass
        # Luôn lấy link mới từ ô nhập, xóa sạch dữ liệu cũ
        self.sheet_url = current_url
        self.num_ban = self.ban_var.get() if hasattr(self, 'ban_var') else None
        self.sheet_rows = []
        creds_path = getattr(self, 'creds_path', None)
        if not creds_path or not os.path.exists(creds_path):
            fallback_creds = [
                os.path.join(os.getcwd(), 'credentials.json'),
                os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'credentials.json')),
            ]
            for cp in fallback_creds:
                if cp and os.path.exists(cp):
                    creds_path = cp
                    self.creds_path = cp
                    try:
                        if hasattr(self, 'creds_label') and self.creds_label:
                            self.creds_label.config(text=os.path.basename(cp), fg='#00FF00')
                    except Exception:
                        pass
                    break
        fetch_matches_from_sheet._creds_path = creds_path
        if self.sheet_url:
            # Always load full source rows; num_ban only controls visible UI rows.
            sheet_rows = fetch_matches_from_sheet(self.sheet_url, None)
            print(f"DEBUG: sheet_rows[0] = {sheet_rows[0] if sheet_rows and not isinstance(sheet_rows, dict) else 'None'}", file=sys.stderr)
            if isinstance(sheet_rows, dict) and 'error' in sheet_rows:
                self.sheet_rows = []
                self.status_var.set(sheet_rows['error'])
                try:
                    with open('error_gsheet.log', 'a', encoding='utf-8') as f:
                        import datetime
                        f.write(f"[{datetime.datetime.now()}] {sheet_rows['error']}\n")
                except Exception:
                    pass
                try:
                    import tkinter as tk
                    from tkinter import messagebox
                    root = None
                    if not hasattr(self, 'winfo_exists') or not self.winfo_exists():
                        root = tk.Tk()
                        root.withdraw()
                    messagebox.showerror('Google Sheet Error', sheet_rows['error'])
                    if root:
                        root.destroy()
                except Exception:
                    pass
                try:
                    print(f"Google Sheet Error: {sheet_rows['error']}", file=sys.stderr)
                except Exception:
                    pass
            elif sheet_rows:
                # Luôn cập nhật dữ liệu mới, xóa sạch dữ liệu cũ
                self.sheet_rows = list(sheet_rows)  # copy mới hoàn toàn
                print(f"DEBUG: self.sheet_rows[0] = {self.sheet_rows[0] if self.sheet_rows else 'None'}", file=sys.stderr)
                self.status_var.set(f'Đã tải {len(sheet_rows)} dòng từ Google Sheet.')
            else:
                self.sheet_rows = []
                self.status_var.set('Không có dữ liệu Google Sheet!')
        else:
            self.sheet_rows = []
            self.status_var.set('Chưa nhập link Google Sheet!')
        # Chỉ xóa bảng và điền dữ liệu mới nếu thực sự có dữ liệu mới từ Google Sheet
        if self.sheet_rows:
            print(f"DEBUG: UI will use self.sheet_rows[0] = {self.sheet_rows[0] if self.sheet_rows else 'None'}", file=sys.stderr)
            self.populate_table()
        else:
            self.status_var.set('Không có dữ liệu Google Sheet! (Bảng cũ được giữ nguyên)')

    def push_to_vmix(self, idx):
        import requests, sys
        widgets = self.match_rows[idx]
        entry_tran = widgets[0]
        tran_raw = entry_tran.get().strip()
        vmix_url = widgets[5].get().strip()
        ban = widgets[1].get()
        vdv_a = widgets[2].get()
        vdv_b = widgets[3].get()
        diem = widgets[4].get()
        # Chỉ gửi dữ liệu lên vMix, không cập nhật Google Sheet
        entry_tran._last_sent = tran_raw
        if hasattr(self, 'row_states') and idx < len(self.row_states):
            self.row_states[idx]['dirty'] = False
            btn = self.match_rows[idx][6]
            btn.config(bg='#388E3C', fg='white')
        tran_fmt = f'Trận {tran_raw}' if tran_raw and not str(tran_raw).startswith('Trận') else str(tran_raw)
        # normalize vmix_url and API base
        try:
            base = vmix_url.strip()
            if not base:
                raise ValueError('Empty vMix URL')
            if not base.startswith('http://') and not base.startswith('https://'):
                base = 'http://' + base
            api_base = base.rstrip('/') + '/API/'
        except Exception as e:
            self.status_var.set(f'Lỗi gửi: không hợp lệ địa chỉ vMix ({e})')
            return

        def _set_text(input_name, selected_name, value):
            try:
                params = {'Function': 'SetText', 'Input': input_name, 'SelectedName': selected_name, 'Value': value}
                requests.get(api_base, params=params, timeout=3)
            except Exception as ee:
                raise

        try:
            # Input 1 fields
            _set_text('1', 'TenA.Text', vdv_a)
            _set_text('1', 'TenB.Text', vdv_b)
            _set_text('1', 'Tran.Text', tran_fmt)
            _set_text('1', 'Noi dung.Text', diem)
            # Backdrop fields (named inputs)
            _set_text('backdrop.gtzip', 'tieu de.Text', self.tengiai_var.get())
            _set_text('backdrop.gtzip', 'Ten A.Text', vdv_a)
            _set_text('backdrop.gtzip', 'Ten B.Text', vdv_b)
            _set_text('backdrop.gtzip', 'thoi gian.Text', self.thoigian_var.get())
            _set_text('backdrop.gtzip', 'dia diem.Text', self.diadiem_var.get())

            # ket qua and ticker fields
            _set_text('ket qua.gtzip', 'tieu de.Text', self.tengiai_var.get())
            _set_text('chay chu', 'Ticker1.Text', self.chuchay_var.get())

            self.status_var.set(f'Đã gửi {tran_fmt} ({ban}) lên vMix: {base}')
        except Exception as e:
            print(f'ERROR: Không gửi được dữ liệu lên vMix: {e}', file=sys.stderr)
            try:
                # Best-effort fallback: try to send minimal fields
                try:
                    _set_text('backdrop.gtzip', 'noi dung.Text', diem)
                except Exception:
                    pass
                try:
                    _set_text('ket qua.gtzip', 'tieu de.Text', self.tengiai_var.get())
                except Exception:
                    pass
                try:
                    _set_text('chay chu', 'Ticker1.Text', self.chuchay_var.get())
                except Exception:
                    pass
            except Exception:
                pass
            friendly = self._friendly_network_error(e, base)
            if tran_fmt:
                self.status_var.set(f'Lỗi gửi {tran_fmt}: {friendly}')
            else:
                self.status_var.set(f'Lỗi gửi: {friendly}')


    def save_table_to_csv(self):
        try:
            if hasattr(self, 'diemso_var') and hasattr(self, 'diemso_text'):
                self.diemso_var.set(self.diemso_text.get('1.0', 'end-1c'))
        except Exception:
            pass

        import csv
        import json
        from tkinter import filedialog

        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            title='Lưu bảng ra file CSV'
        )
        if not file_path:
            return

        try:
            try:
                preview_win = getattr(self, '_preview_window', None)
                if preview_win is not None and getattr(preview_win, 'winfo_exists', lambda: False)():
                    self._update_last_preview_meta(preview_win)
            except Exception:
                pass

            serial_preview = getattr(self, '_last_preview_meta', None)
            try:
                preview_json = json.dumps(serial_preview, ensure_ascii=False)
            except Exception:
                preview_json = ''

            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, quoting=csv.QUOTE_ALL)
                writer.writerow(['# TenGiai', self.tengiai_var.get() if hasattr(self, 'tengiai_var') else ''])
                writer.writerow(['# ChuChay', self.chuchay_var.get() if hasattr(self, 'chuchay_var') else ''])
                writer.writerow(['# ThoiGian', self.thoigian_var.get() if hasattr(self, 'thoigian_var') else ''])
                writer.writerow(['# DiaDiem', self.diadiem_var.get() if hasattr(self, 'diadiem_var') else ''])
                writer.writerow(['# DiemSo', self.diemso_text.get('1.0', 'end-1c') if hasattr(self, 'diemso_text') else ''])
                writer.writerow(['# LinkGoogleSheet', self.url_var.get().strip() if hasattr(self, 'url_var') else ''])
                writer.writerow(['# CredentialsPath', self.creds_path if getattr(self, 'creds_path', None) else ''])
                writer.writerow(['# HbsfUrl', self.hbsf_url_var.get().strip() if hasattr(self, 'hbsf_url_var') else 'https://hbsf.com.vn'])
                writer.writerow(['# EventId', self.event_id_var.get().strip() if hasattr(self, 'event_id_var') else ''])
                writer.writerow(['# RoundType', self.round_type_var.get() if hasattr(self, 'round_type_var') else 'Vòng Loại'])
                writer.writerow(['# TotalTables', self.ban_var.get() if hasattr(self, 'ban_var') and self.ban_var is not None else 0])
                writer.writerow(['# PreviewConfig', preview_json])
                writer.writerow(['# PreviewFooterLogo', str(getattr(self, '_preview_footer_logo_path', '') or '')])
                writer.writerow(['# DateSaved', __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                writer.writerow(['Số trận', 'Số bàn', 'Tên VĐV A', 'Tên VĐV B', 'Điểm số', 'Địa chỉ vMix'])

                for row in getattr(self, 'match_rows', []) or []:
                    values = [
                        row[0].get() if len(row) > 0 and hasattr(row[0], 'get') else '',
                        row[1].get() if len(row) > 1 and hasattr(row[1], 'get') else '',
                        row[2].get() if len(row) > 2 and hasattr(row[2], 'get') else '',
                        row[3].get() if len(row) > 3 and hasattr(row[3], 'get') else '',
                        row[4].get() if len(row) > 4 and hasattr(row[4], 'get') else '',
                        row[5].get() if len(row) > 5 and hasattr(row[5], 'get') else '',
                    ]
                    values = values[:6] + [''] * (6 - len(values))
                    writer.writerow(values)

            try:
                self._auto_save_state()
            except Exception:
                pass
            self.status_var.set(f'Đã lưu bảng ra {file_path} (kèm cấu hình Preview)')
        except Exception as e:
            self.status_var.set(f'Lỗi lưu CSV: {e}')

    def load_table_from_csv(self):
        import csv
        import json
        from tkinter import filedialog, messagebox

        file_path = filedialog.askopenfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            title='Tải bảng từ file CSV'
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
        except Exception as ex:
            self.status_var.set(f'Lỗi đọc CSV: {ex}')
            return

        total_tables = None
        diemso_loaded = ''
        preview_loaded = None
        preview_footer_logo_loaded = ''
        table_start = None

        for idx, row in enumerate(rows):
            if not row:
                continue

            first_col = str(row[0]).strip().replace('\ufeff', '')
            val = row[1] if len(row) > 1 else ''

            if first_col.startswith('#'):
                key = first_col[1:].strip().lower()
                if key.startswith('tengiai') and hasattr(self, 'tengiai_var'):
                    self.tengiai_var.set(val)
                elif key.startswith('chuchay') and hasattr(self, 'chuchay_var'):
                    self.chuchay_var.set(val)
                elif key.startswith('thoigian') and hasattr(self, 'thoigian_var'):
                    self.thoigian_var.set(val)
                elif key.startswith('diadiem') and hasattr(self, 'diadiem_var'):
                    self.diadiem_var.set(val)
                elif key.startswith('diemso'):
                    diemso_loaded = val
                elif key.startswith('linkgooglesheet') and hasattr(self, 'url_var'):
                    self.url_var.set(val)
                elif key.startswith('credentialspath'):
                    self.creds_path = val
                    try:
                        self.creds_label.config(
                            text=self.creds_path if self.creds_path else '(Chưa chọn credentials)',
                            fg='#00FF00' if self.creds_path else '#FFD369'
                        )
                    except Exception:
                        pass
                elif key.startswith('hbsfurl'):
                    if hasattr(self, 'hbsf_url_var'):
                        self.hbsf_url_var.set(val)
                    else:
                        self._saved_hbsf_url = val
                elif key.startswith('eventid'):
                    if hasattr(self, 'event_id_var'):
                        self.event_id_var.set(val)
                    else:
                        self._saved_event_id = val
                elif key.startswith('roundtype'):
                    if hasattr(self, 'round_type_var'):
                        self.round_type_var.set(val)
                    else:
                        self._saved_round_type = val
                elif key.startswith('totaltables'):
                    try:
                        total_tables = int(val)
                    except Exception:
                        total_tables = None
                elif key.startswith('previewconfig'):
                    try:
                        parsed = json.loads(val) if val else None
                        if isinstance(parsed, list):
                            normalized = []
                            for i in range(9):
                                cell = parsed[i] if i < len(parsed) else None
                                if isinstance(cell, dict):
                                    normalized.append({
                                        'type': cell.get('type'),
                                        'value': cell.get('value'),
                                        'image_mode': cell.get('image_mode', 'fit'),
                                        'logo_effect': cell.get('logo_effect', 'cut'),
                                        'logo_interval': cell.get('logo_interval', 4.0),
                                    })
                                else:
                                    normalized.append(None)
                            preview_loaded = normalized
                    except Exception:
                        preview_loaded = None
                elif key.startswith('previewfooterlogo'):
                    preview_footer_logo_loaded = str(val or '').strip()
                continue

            first_norm = first_col.lower()
            if first_norm == 'số trận' and len(row) >= 6:
                table_start = idx
                break

        if table_start is None:
            preview_text = '\n'.join([','.join(r) for r in rows[:20]])
            messagebox.showerror(
                'File CSV không hợp lệ',
                'Không tìm thấy tiêu đề bảng hợp lệ trong file CSV này.\n\n'
                f'Nội dung file (20 dòng đầu):\n---\n{preview_text}\n---\n\n'
                'Hãy kiểm tra lại định dạng, dấu phẩy, ký tự lạ hoặc gửi file này cho kỹ thuật viên để hỗ trợ.'
            )
            return

        try:
            self.clear_table()
        except Exception:
            pass

        if total_tables is not None and hasattr(self, 'ban_var') and self.ban_var is not None:
            try:
                self.ban_var.set(total_tables)
            except Exception:
                pass

        try:
            self.populate_table()
        except Exception:
            pass

        if hasattr(self, 'diemso_text'):
            try:
                self.diemso_text.delete('1.0', 'end')
                self.diemso_text.insert('1.0', diemso_loaded)
            except Exception:
                pass

        data_rows = rows[table_start + 1: table_start + 1 + len(getattr(self, 'match_rows', []) or [])]
        for ridx, csv_row in enumerate(data_rows):
            values = list(csv_row) if csv_row else []
            values = values[:6] + [''] * (6 - len(values))
            if not any(str(cell).strip() for cell in values):
                continue

            if ridx >= len(getattr(self, 'match_rows', []) or []):
                continue
            for col in range(6):
                try:
                    widget = self.match_rows[ridx][col]
                    widget.config(state='normal')
                    widget.delete(0, 'end')
                    widget.insert(0, values[col])
                    if col in [2, 3]:
                        widget.config(state='readonly')
                except Exception:
                    continue

        if isinstance(preview_loaded, list):
            try:
                self._last_preview_meta = preview_loaded
            except Exception:
                pass

            preview_window = getattr(self, '_preview_window', None)
            if preview_window is not None and getattr(preview_window, 'winfo_exists', lambda: False)():
                try:
                    new_meta = []
                    for cell in preview_loaded:
                        if isinstance(cell, dict):
                            new_meta.append({
                                'type': cell.get('type'),
                                'value': cell.get('value'),
                                'image_ref': None,
                                'image_mode': cell.get('image_mode', 'fit'),
                                'logo_effect': cell.get('logo_effect', 'cut'),
                                'logo_interval': cell.get('logo_interval', 4.0),
                            })
                        else:
                            new_meta.append({'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'})
                    preview_window.cell_meta = new_meta
                    for frame in getattr(preview_window, 'cells', []) or []:
                        try:
                            frame.event_generate('<Configure>')
                        except Exception:
                            pass
                except Exception:
                    pass

        try:
            self._preview_footer_logo_path = preview_footer_logo_loaded
        except Exception:
            pass
        try:
            self._apply_preview_footer_logo_to_open_window()
        except Exception:
            pass

        try:
            self._auto_save_state()
        except Exception:
            pass
        self.status_var.set(f'Đã tải bảng từ {file_path} (kèm cấu hình Preview nếu có)')

if __name__ == '__main__':
    print('DEBUG: Starting FullScreenMatchGUI...')
    app = FullScreenMatchGUI()
    print('DEBUG: Entering mainloop...')
    app.mainloop()
    print('DEBUG: mainloop exited.')
