import sys, os
import atexit
import tkinter as tk
from tkinter import ttk, messagebox


# Helper to fetch rows from Google Sheets using GSheetClient
_GSheetClient = None
try:
    from src.gsheet_client import GSheetClient as _GSheetClient
except Exception:
    try:
        from gsheet_client import GSheetClient as _GSheetClient
    except Exception:
        _GSheetClient = None

# Public alias used by the module
GSheetClient = _GSheetClient

def fetch_matches_from_sheet(sheet_url: str, max_rows: int = None):
    """Return list of row dicts from 'Kết quả' sheet or {'error': msg} on failure.
    Uses `fetch_matches_from_sheet._creds_path` (may be set by GUI) for credentials.
    """
    try:
        if not sheet_url:
            return []
        import re
        m = re.search(r"/spreadsheets/d/([\w-]+)", sheet_url)
        spreadsheet_id = m.group(1) if m else None
        creds = getattr(fetch_matches_from_sheet, '_creds_path', None)
        if not (spreadsheet_id and creds and os.path.exists(creds)):
            return {'error': 'Missing spreadsheet ID or credentials (choose credentials)'}
        # Prefer the globally imported GSheetClient if available
        gs_client_cls = globals().get('GSheetClient', None)
        if gs_client_cls is None:
            # Try to import from local src path dynamically
            try:
                # add the package dir to sys.path
                base = os.path.dirname(os.path.abspath(__file__))
                if base not in sys.path:
                    sys.path.insert(0, base)
                from gsheet_client import GSheetClient as gs_client_cls
            except Exception as ie:
                return {'error': f'Cannot import GSheetClient: {ie}'}
        else:
            gs_client_cls = gs_client_cls
        try:
            gs = gs_client_cls(spreadsheet_id, creds)
        except Exception as ex:
            return {'error': str(ex)}
        rng = 'Kết quả!A1:Z2000'
        rows = gs.read_table(rng) or []
        return rows
    except Exception as ex:
        return {'error': str(ex)}


def fetch_matches_from_ketqua(sheet_url: str, max_rows: int = None):
    return fetch_matches_from_sheet(sheet_url, max_rows)
""" Removed duplicate `render_cell` block that was accidentally inserted
at the top of the file and caused IndentationError. The real
`render_cell` is defined inside `open_preview_all()` later in the file.
                frame = preview.cells[idx]
                # clear (but keep a small control frame reserved)
                for w in frame.winfo_children():
                    w.destroy()
                meta = preview.cell_meta[idx]

                # Overlay control (fixed small corner button)
                ctrl = tk.Frame(frame, bg='#111')
                ctrl.place(relx=1.0, rely=0.0, anchor='ne')
                tk.Button(ctrl, text='⋮', width=2, command=lambda i=idx: config_cell(i), bg='#FFD369').pack()

                # Determine available drawing area (exclude some padding)
                fw = max(10, frame.winfo_width())
                fh = max(10, frame.winfo_height())
                pad = 12
                avail_w = max(10, fw - pad)
                avail_h = max(10, fh - pad)

                # Helper to place scaled text labels that try to fill the cell
                def place_scoreboard(ten_a, ten_b, tran, diem_a, diem_b, avg_a, avg_b, lco):
                    # Estimate font sizes based on available height
                    # Reserve ~55% height for main scores, rest for names/sub
                    main_h = int(avail_h * 0.55)
                    name_h = int(avail_h * 0.18)
                    sub_h = max(10, avail_h - main_h - name_h - 8)

                    # Determine approximate font sizes
                    # Using pixel sizes with tkinter Font requires computing from points; approximate by dividing height
                    score_font_size = max(12, main_h // 2)
                    name_font_size = max(10, name_h // 2)
                    sub_font_size = max(9, sub_h // 3)

                    # Build layout
                    top = tk.Frame(frame, bg='#000')
                    top.place(relx=0.5, rely=0.05, anchor='n')
                    tk.Label(top, text=ten_a, font=('Arial', name_font_size, 'bold'), fg='white', bg='#000').pack()

                    mid = tk.Frame(frame, bg='#000')
                    mid.place(relx=0.5, rely=0.3, anchor='n')
                    # Scores side by side
                    left = tk.Label(mid, text=str(diem_a or '-'), font=('Arial', score_font_size, 'bold'), fg='white', bg='#000')
                    right = tk.Label(mid, text=str(diem_b or '-'), font=('Arial', score_font_size, 'bold'), fg='#FFD600', bg='#000')
                    left.pack(side='left', padx=20)
                    tk.Label(mid, text=f'TRẬN {tran}', font=('Arial', sub_font_size, 'bold'), fg='#00E5FF', bg='#000').pack(side='left', padx=8)
                    right.pack(side='left', padx=20)

                    btm = tk.Frame(frame, bg='#000')
                    btm.place(relx=0.5, rely=0.75, anchor='n')
                    tk.Label(btm, text=f'AVG {avg_a}', fg='white', bg='#000', font=('Arial', sub_font_size)).pack(side='left', padx=6)
                    tk.Label(btm, text=f'Lco {lco}', fg='#FFD369', bg='#000', font=('Arial', sub_font_size)).pack(side='left', padx=6)
                    tk.Label(btm, text=f'AVG {avg_b}', fg='#FFD600', bg='#000', font=('Arial', sub_font_size)).pack(side='left', padx=6)

                if meta['type'] == 'vmix' and meta['value']:
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
                            # Use place_scoreboard to draw scaled labels
                            place_scoreboard(ten_a, ten_b, tran, diem_a, diem_b, avg_a, avg_b, lco)
                        else:
                            tk.Label(frame, text='No vMix URL', fg='white', bg='#000').place(relx=0.5, rely=0.5, anchor='center')
                    except Exception as ex:
                        tk.Label(frame, text=f'Error fetch vMix:\n{ex}', fg='red', bg='#000').place(relx=0.5, rely=0.5, anchor='center')
                elif meta.get('type') == 'row_field' and meta.get('value') is not None:
                    # meta['value'] is (rowidx, field_key)
                    try:
                        rowidx, field = meta.get('value')
                        if isinstance(rowidx, int) and rowidx < len(self.match_rows):
                            row = self.match_rows[rowidx]
                            # map field keys to column indices or special handling
                            field_map = {
                                'match': 0,
                                'table': 1,
                                'name_a': 2,
                                'name_b': 3,
                                'score': 4,
                                'vmix': 5,
                            }
                            text = ''
                            if field in field_map:
                                try:
                                    widget = row[field_map[field]]
                                    text = widget.get() if hasattr(widget, 'get') else ''
                                except Exception:
                                    text = ''
                            else:
                                text = str(field)
                        else:
                            text = ''
                        fw, fh = cell_size(frame)
                        # scale font size to cell
                        try:
                            size = max(10, min(72, int(min(fw, fh) // 12)))
                        except Exception:
                            size = 16
                        lbl = tk.Label(frame, text=text, font=('Arial', size, 'bold'), fg='white', bg='#000', wraplength=int(fw*0.9), justify='center')
                        lbl.pack(expand=True)
                    except Exception as ex:
                        tk.Label(frame, text=f'Error render field:\n{ex}', fg='red', bg='#000').pack(padx=8, pady=8)
                elif meta['type'] == 'image' and meta['value']:
                    path = meta['value']
                    mode = meta.get('image_mode', 'fit')
                    if Image and ImageTk:
                        try:
                            img = Image.open(path)
                            iw, ih = img.size
                            # Compute scaling
                            if mode == 'fit':
                                # contain: scale to fit within avail_w x avail_h
                                scale = min(avail_w / iw, avail_h / ih)
                            elif mode == 'cover':
                                # cover: scale to fill and crop
                                scale = max(avail_w / iw, avail_h / ih)
                            else:
                                # center: no scale beyond fitting, show centered
                                scale = min(1.0, min(avail_w / iw, avail_h / ih))
                            new_w = max(1, int(iw * scale))
                            new_h = max(1, int(ih * scale))
                            img2 = img.resize((new_w, new_h), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
                            # For cover, crop center to avail_w x avail_h
                            if mode == 'cover':
                                left = max(0, (new_w - avail_w)//2)
                                upper = max(0, (new_h - avail_h)//2)
                                right = left + avail_w
                                lower = upper + avail_h
                                img2 = img2.crop((left, upper, right, lower))
                                tkimg = ImageTk.PhotoImage(img2)
                                lbl = tk.Label(frame, image=tkimg, bg='#000')
                                lbl.image = tkimg
                                lbl.place(relx=0.5, rely=0.5, anchor='center')
                            else:
                                tkimg = ImageTk.PhotoImage(img2)
                                lbl = tk.Label(frame, image=tkimg, bg='#000')
                                lbl.image = tkimg
                                # center the image
                                lbl.place(relx=0.5, rely=0.5, anchor='center')
                            preview.cell_meta[idx]['image_ref'] = tkimg
                        except Exception as ex:
                            tk.Label(frame, text=f'Image error:\n{ex}', fg='red', bg='#000').place(relx=0.5, rely=0.5, anchor='center')
                    else:
                        tk.Label(frame, text='Pillow not installed\nCannot display image', fg='white', bg='#000').place(relx=0.5, rely=0.5, anchor='center')
                else:
                    # empty placeholder / poster
                    tk.Label(frame, text='Poster / Logo', fg='#FFD369', bg='#000', font=('Arial', 20, 'bold')).place(relx=0.5, rely=0.5, anchor='center')
"""

class FullScreenMatchGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        # ensure some commonly-used instance attributes exist before other init steps
        try:
            self.match_rows = []
        except Exception:
            pass
        
        try:
            self._auto_state_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ui_state.pkl')
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
        # Attempt to eagerly load saved state so restore runs during widget init
        try:
            import pickle, sys
            candidates = []
            if getattr(self, '_auto_state_path', None):
                candidates.append(self._auto_state_path)
            candidates.append(os.path.join(os.getcwd(), 'ui_state.pkl'))
            candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ui_state.pkl'))
            try:
                exe_dir = os.path.dirname(sys.executable)
                candidates.append(os.path.join(exe_dir, 'ui_state.pkl'))
            except Exception:
                pass
            for p in candidates:
                try:
                    if p and os.path.exists(p):
                        with open(p, 'rb') as f:
                            s = pickle.load(f)
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
                    try:
                        self.write_all_vmix_to_sheet()
                    except Exception:
                        pass
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
            import os
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
            pw, ph = 760, 720
            popup.geometry(f'{pw}x{ph}')
            try:
                popup.minsize(pw, ph)
                popup.resizable(True, True)
            except Exception:
                pass
            popup.configure(bg='#1A2233')
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
                popup.geometry(f'{pw}x{ph}+{x}+{y}')
            except Exception:
                pass
            entries = {}
            lbl_fg = '#FFD369'
            entry_bg = '#232B3E'
            entry_fg = 'white'
            fnt_label = ('Segoe UI', 14, 'bold')
            fnt_entry = ('Segoe UI', 14)

            # Column headings (A / B)
            tk.Label(popup, text='VĐV A', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=0, column=0, padx=12, pady=(12,6), sticky='w')
            tk.Label(popup, text='VĐV B', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=0, column=1, padx=12, pady=(12,6), sticky='w')

            # Tên
            tk.Label(popup, text='Tên', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=1, column=0, sticky='w', padx=12)
            e_tena = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_tena.grid(row=2, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Tên VĐV A'] = e_tena

            tk.Label(popup, text='Tên', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=1, column=1, sticky='w', padx=12)
            e_tenb = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_tenb.grid(row=2, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Tên VĐV B'] = e_tenb

            # Điểm row
            tk.Label(popup, text='Điểm', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=3, column=0, sticky='w', padx=12)
            e_diem_a = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_diem_a.grid(row=4, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Điểm A'] = e_diem_a

            tk.Label(popup, text='Điểm', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=3, column=1, sticky='w', padx=12)
            e_diem_b = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_diem_b.grid(row=4, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Điểm B'] = e_diem_b

            # Lượt cơ should appear under the Điểm row, spanning both columns
            tk.Label(popup, text='Lượt cơ', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=5, column=0, columnspan=2, sticky='w', padx=12)
            e_lco = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_lco.grid(row=6, column=0, columnspan=2, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['Lượt cơ'] = e_lco

            # HR rows: HR1 and HR2 for A and B
            tk.Label(popup, text='HR1', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=7, column=0, sticky='w', padx=12)
            e_hr1a = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_hr1a.grid(row=8, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR1A'] = e_hr1a

            tk.Label(popup, text='HR1', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=7, column=1, sticky='w', padx=12)
            e_hr1b = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_hr1b.grid(row=8, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR1B'] = e_hr1b

            tk.Label(popup, text='HR2', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=9, column=0, sticky='w', padx=12)
            e_hr2a = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_hr2a.grid(row=10, column=0, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR2A'] = e_hr2a

            tk.Label(popup, text='HR2', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=9, column=1, sticky='w', padx=12)
            e_hr2b = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry)
            e_hr2b.grid(row=10, column=1, padx=12, pady=(0,8), sticky='ew', ipady=6)
            entries['HR2B'] = e_hr2b

            # AVGs
            tk.Label(popup, text='AVGA', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=11, column=0, sticky='w', padx=12)
            e_avga = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry, state='normal')
            e_avga.grid(row=12, column=0, padx=12, pady=(0,12), sticky='ew', ipady=6)
            entries['AVGA'] = e_avga

            tk.Label(popup, text='AVGB', fg=lbl_fg, bg='#1A2233', font=fnt_label).grid(row=11, column=1, sticky='w', padx=12)
            e_avgb = tk.Entry(popup, bg=entry_bg, fg=entry_fg, insertbackground=entry_fg, font=fnt_entry, state='normal')
            e_avgb.grid(row=12, column=1, padx=12, pady=(0,12), sticky='ew', ipady=6)
            entries['AVGB'] = e_avgb

            # make grid expand entries
            popup.grid_columnconfigure(0, weight=1)
            popup.grid_columnconfigure(1, weight=1)

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
                        'Tên VĐV A': ['TenA','TenA.Text'],'Tên VĐV B': ['TenB','TenB.Text']
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
                        return
                    mapping = {
                        'Tên VĐV A': 'TenA.Text','Tên VĐV B':'TenB.Text',
                        'Điểm A':'DiemA.Text','Điểm B':'DiemB.Text','Lượt cơ':'Lco.Text',
                        'HR1A':'HR1A.Text','HR2A':'HR2A.Text','HR1B':'HR1B.Text','HR2B':'HR2B.Text',
                        'AVGA':'AvgA.Text','AVGB':'AvgB.Text'
                    }
                    session = requests.Session()
                    for fld, sel in mapping.items():
                        val = entries[fld].get() if entries.get(fld) else ''
                        try:
                            session.get(f"{vmix_url}/API/", params={'Function':'SetText','Input':1,'SelectedName':sel,'Value':val}, timeout=2)
                        except Exception:
                            pass
                    messagebox.showinfo('Thành công', 'Đã gửi lên vMix')
                except Exception as ex:
                    messagebox.showerror('Lỗi', f'Không gửi được: {ex}')

            # Buttons (aligned at bottom)
            btn_frame = tk.Frame(popup, bg='#1A2233')
            btn_frame.grid(row=13, column=0, columnspan=2, pady=(8,12))
            tk.Button(btn_frame, text='Lấy từ vMix', command=prefill_from_vmix, bg='#2962FF', fg='#FFD369', font=('Segoe UI', 13, 'bold')).pack(side='left', padx=8, ipadx=10, ipady=6)
            # send and close helper binds Enter and button
            def send_and_close(event=None):
                try:
                    send_to_vmix()
                finally:
                    try:
                        popup.destroy()
                    except Exception:
                        pass

            tk.Button(btn_frame, text='Gửi', command=send_and_close, bg='#00C853', fg='#1A2233', font=('Segoe UI', 13, 'bold')).pack(side='left', padx=8, ipadx=10, ipady=6)
            tk.Button(btn_frame, text='Hủy', command=popup.destroy, bg='#FF6F00', fg='white', font=('Segoe UI', 13, 'bold')).pack(side='left', padx=8, ipadx=10, ipady=6)

            # bind Enter to send
            try:
                popup.bind('<Return>', send_and_close)
                popup.bind('<KP_Enter>', send_and_close)
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
        """Gán lại giá trị vào UI từ self._restored_state nếu có"""
        # If no restored state is loaded yet, attempt to read the auto-save pickle file.
        s = getattr(self, '_restored_state', None)
        if not s:
            # Try loading from file
            # Try multiple candidate locations for ui_state.pkl to be robust
            try:
                import pickle, sys
                candidates = []
                if getattr(self, '_auto_state_path', None):
                    candidates.append(self._auto_state_path)
                # cwd
                candidates.append(os.path.join(os.getcwd(), 'ui_state.pkl'))
                # module dir
                candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ui_state.pkl'))
                # executable dir (useful for bundled exe)
                try:
                    exe_dir = os.path.dirname(sys.executable)
                    candidates.append(os.path.join(exe_dir, 'ui_state.pkl'))
                except Exception:
                    pass
                # try load first readable candidate
                s = None
                for p in candidates:
                    try:
                        # log candidate check
                        try:
                            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                            import datetime
                            with open(dbg, 'a', encoding='utf-8') as df:
                                df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE_CANDIDATE check={p} exists={os.path.exists(p)}\n")
                        except Exception:
                            pass
                        if p and os.path.exists(p):
                            with open(p, 'rb') as f:
                                s = pickle.load(f)
                                self._restored_state = s
                                # preserve preview snapshot in-memory so autosave won't clear it
                                try:
                                    if s.get('preview'):
                                        self._last_preview_meta = s.get('preview')
                                except Exception:
                                    pass
                                try:
                                    dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                                    import datetime
                                    with open(dbg, 'a', encoding='utf-8') as df:
                                        df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE_LOAD path={p} keys={list(s.keys())} table_rows={len(s.get('table',[]))} preview_present={bool(s.get('preview'))}\n")
                                except Exception:
                                    pass
                                break
                    except Exception:
                        continue
            except Exception:
                s = None
        if not s:
            return
        # mark restore as in-progress: prevent autosave writes until restore is committed
        try:
            self._restore_committed = False
        except Exception:
            pass
        # IMMEDIATELY suspend autosave during restore to prevent overwriting restored state
        try:
            import time
            self._autosave_suspended_until = time.time() + 5.5
        except Exception:
            pass
        # Log that we restored state (helps debugging headless runs)
        try:
            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
            import datetime
            with open(dbg, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now().isoformat()}] AUTO_RESTORE_STATE loaded table_rows={len(s.get('table', []))}\n")
        except Exception:
            pass
        # Apply restored state to UI widgets if they exist
        try:
            # global fields
            if hasattr(self, 'tengiai_var') and s.get('tengiai') is not None:
                try:
                    self.tengiai_var.set(s.get('tengiai',''))
                except Exception:
                    pass
            if hasattr(self, 'thoigian_var') and s.get('thoigian') is not None:
                try:
                    self.thoigian_var.set(s.get('thoigian',''))
                except Exception:
                    pass
            if hasattr(self, 'diadiem_var') and s.get('diadiem') is not None:
                try:
                    self.diadiem_var.set(s.get('diadiem',''))
                except Exception:
                    pass
            if hasattr(self, 'chuchay_var') and s.get('chuchay') is not None:
                try:
                    self.chuchay_var.set(s.get('chuchay',''))
                except Exception:
                    pass
            if hasattr(self, 'chuchay_text') and s.get('chuchay') is not None:
                try:
                    self.chuchay_text.delete('1.0', 'end')
                    self.chuchay_text.insert('1.0', s.get('chuchay',''))
                except Exception:
                    pass
            if hasattr(self, 'diemso_text') and s.get('diemso') is not None:
                try:
                    val = s.get('diemso','')
                    self.diemso_text.delete('1.0', 'end')
                    self.diemso_text.insert('1.0', val)
                    # propagate master score into table rows
                    try:
                        self.propagate_master_score(val)
                    except Exception:
                        pass
                except Exception:
                    pass
            if hasattr(self, 'url_var') and s.get('sheet_url') is not None:
                try:
                    self.url_var.set(s.get('sheet_url',''))
                except Exception:
                    pass
            if s.get('creds_path'):
                try:
                    self.creds_path = s.get('creds_path')
                    if hasattr(self, 'creds_label'):
                        self.creds_label.config(text=os.path.basename(self.creds_path))
                except Exception:
                    pass
            # Restore tổng số bàn trước (luôn set, kể cả khi là 0)
            try:
                ban_val = s.get('ban')
                # Nếu không có trường ban, lấy số dòng của bảng (table)
                if ban_val is None:
                    ban_val = len(s.get('table', []))
                # Ghi log khi khôi phục số bàn
                try:
                    dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                    import datetime
                    with open(dbg, 'a', encoding='utf-8') as df:
                        df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE ban field: {ban_val}\n")
                except Exception:
                    pass
                if ban_val is not None:
                    self.ban_var.set(ban_val)
            except Exception:
                pass
            self._restoring_state = True
            # Tạo bảng đúng số dòng
            try:
                self.populate_table()
            except Exception:
                pass
            # Gán lại giá trị từng dòng (bao gồm Địa chỉ vMix, điểm số, tên, v.v.)
            table = s.get('table', [])
            if table:
                try:
                    self._apply_restored_table(table)
                except Exception:
                    pass
            self._restoring_state = False
            # Set lại điểm số tổng và propagate xuống từng dòng
            if hasattr(self, 'diemso_text') and s.get('diemso') is not None:
                try:
                    val = s.get('diemso','')
                    self.diemso_text.delete('1.0', 'end')
                    self.diemso_text.insert('1.0', val)
                    self.propagate_master_score(val)
                except Exception:
                    pass
            # restore preview configuration (if saved)
            try:
                preview_data = s.get('preview', None)
                if preview_data:
                    # open preview window (creates frames)
                    try:
                        p = self.preview_open()
                    except Exception:
                        p = getattr(self, '_preview_window', None)
                    if p is not None:
                        # rebuild cell_meta from serializable data
                        new_meta = []
                        for cell in preview_data:
                            try:
                                if cell is None:
                                    new_meta.append({'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'})
                                else:
                                    new_meta.append({'type': cell.get('type'), 'value': cell.get('value'), 'image_ref': None, 'image_mode': cell.get('image_mode', 'fit')})
                            except Exception:
                                new_meta.append({'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'})
                        try:
                            p.cell_meta = new_meta
                        except Exception:
                            pass
                        # persist as last-known preview snapshot so autosave doesn't lose it
                        try:
                            self._last_preview_meta = []
                            for cell in preview_data:
                                try:
                                    self._last_preview_meta.append({'type': cell.get('type'), 'value': cell.get('value'), 'image_mode': cell.get('image_mode', 'fit')})
                                except Exception:
                                    self._last_preview_meta.append(None)
                        except Exception:
                            pass
                        # briefly suspend autosave to avoid immediate overwrite
                        try:
                            import time
                            self._autosave_suspended_until = time.time() + 2.5
                        except Exception:
                            pass
                        # trigger redraw for each cell by generating <Configure>
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
                    else:
                        # couldn't open preview now; schedule a retry to apply preview when UI ready
                        try:
                            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                            import datetime
                            with open(dbg, 'a', encoding='utf-8') as df:
                                df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE_PREVIEW schedule retry\n")
                        except Exception:
                            pass
                        try:
                            self.after(500, lambda pd=preview_data: self._restore_preview_later(pd))
                        except Exception:
                            pass
            except Exception:
                pass
        except Exception:
            pass

        # After restore completed, give layout callbacks time, then mark restore committed
        try:
            import time
            self._autosave_suspended_until = time.time() + 6.0
        except Exception:
            pass
        try:
            dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
            import datetime
            with open(dbg, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now().isoformat()}] RESTORE_COMPLETE: scheduling final committed save\n")
        except Exception:
            pass
        try:
            # mark restore committed so autosave is allowed and perform a final save
            try:
                self._restore_committed = True
            except Exception:
                pass
            self._auto_save_state()
        except Exception:
            pass

    # --- Programmatic preview control helpers ---
    def preview_open(self):
        """Mở cửa sổ preview (gọi action của nút nếu cần). Trả về đối tượng preview hoặc None."""
        try:
            # If preview already exists and is valid, just deiconify
            p = getattr(self, '_preview_window', None)
            if p is not None:
                try:
                    if p.winfo_exists():
                        try:
                            p.deiconify()
                        except Exception:
                            pass
                        return p
                except Exception:
                    pass
            # Otherwise, invoke the preview button to create it
            try:
                print('[DEBUG] preview_open: preview_all_btn exists?', hasattr(self, 'preview_all_btn'))
                if hasattr(self, 'preview_all_btn'):
                    try:
                        self.preview_all_btn.invoke()
                        print('[DEBUG] preview_open: invoked preview_all_btn')
                    except Exception as e:
                        print('[DEBUG] preview_open: invoke failed', e)
                elif hasattr(self, 'open_preview_all'):
                    try:
                        self.open_preview_all()
                        print('[DEBUG] preview_open: called open_preview_all fallback')
                    except Exception as e:
                        print('[DEBUG] preview_open: fallback open_preview_all failed', e)
            except Exception as e:
                print('[DEBUG] preview_open: unexpected error', e)
            return getattr(self, '_preview_window', None)
        except Exception:
            return None

    def preview_set_cell(self, idx, kind, value, image_mode='fit'):
        """Cấu hình ô preview theo mã.
        kind: 'image' hoặc 'vmix' (string)
        value: đường dẫn ảnh hoặc URL / row index
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
            else:
                meta[idx] = {'type': 'vmix', 'value': value, 'image_ref': None}
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
            from tkinter import filedialog, simpledialog, messagebox
            try:
                from PIL import Image, ImageTk
            except Exception:
                Image = ImageTk = None

            preview = tk.Toplevel(self)
            try:
                self._preview_window = preview
            except Exception:
                pass
            preview.title('Preview 3x3 Bảng điểm')
            try:
                preview.attributes('-fullscreen', True)
            except Exception:
                try:
                    preview.state('zoomed')
                except Exception:
                    pass

            preview.cells = []
            # Auto-map vMix/table rows to preview grid cells if match_rows is available
            num_tables = len(getattr(self, 'match_rows', []))
            # Default: all empty
            preview.cell_meta = [{'type': None, 'value': None, 'image_ref': None, 'image_mode': 'fit'} for _ in range(9)]
            if num_tables in (6, 8):
                # Mapping for 6 tables: [0,1,2,3,4,5] to [0,1,2,3,4,5] (cells 0-5)
                # Mapping for 8 tables: [0,1,2,3,4,5,6,7] to [0,1,2,3,5,6,7,8] (skip cell 4)
                if num_tables == 6:
                    for i in range(6):
                        preview.cell_meta[i] = {'type': 'vmix', 'value': i, 'image_ref': None, 'image_mode': 'fit'}
                elif num_tables == 8:
                    mapping = [0,1,2,3,5,6,7,8]  # table row 0->cell0, 1->1, 2->2, 3->3, 4->5, 5->6, 6->7, 7->8
                    for table_idx, cell_idx in enumerate(mapping):
                        preview.cell_meta[cell_idx] = {'type': 'vmix', 'value': table_idx, 'image_ref': None, 'image_mode': 'fit'}

            def cell_size(frame):
                try:
                    fw = max(10, frame.winfo_width())
                    fh = max(10, frame.winfo_height())
                except Exception:
                    fw = preview.winfo_screenwidth() // 3
                    fh = preview.winfo_screenheight() // 3
                return fw, fh

            def render_cell_local(idx):
                cell = preview.cell_meta[idx]
                frame = preview.cells[idx]
                for w in frame.winfo_children():
                    w.destroy()
                w, h = cell_size(frame)
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
                        else:
                            scale = max(w/iw, h/ih)
                            nw = max(1, int(iw*scale))
                            nh = max(1, int(ih*scale))
                            img2 = img.resize((nw, nh), Image.LANCZOS)
                            left = (nw - w)//2
                            top = (nh - h)//2
                            img3 = img2.crop((left, top, left + w, top + h))
                            tkimg = ImageTk.PhotoImage(img3)
                            lbl = tk.Label(frame, image=tkimg, bg='#000')
                            lbl.image = tkimg
                            lbl.place(x=0, y=0, width=w, height=h)
                    except Exception:
                        tk.Label(frame, text='Image error', fg='white', bg='red').pack(expand=True)
                elif cell.get('type') == 'vmix':
                    try:
                        if isinstance(cell['value'], int):
                            rowidx = cell['value']
                            if rowidx < len(getattr(self, 'match_rows', [])):
                                row = self.match_rows[rowidx]
                                vmix_url = row[5].get().strip() if hasattr(row[5], 'get') else ''
                            else:
                                vmix_url = ''
                        else:
                            vmix_url = str(cell['value'])
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
                            # dynamic layout: compute best font sizes to fill available area
                            from tkinter import font as tkfont
                            name_text = ten_a
                            score_text_a = str(diem_a)
                            score_text_b = str(diem_b)
                            # binary search for largest score font size that fits width/height
                            lo, hi = 8, 300
                            best = 12
                            while lo <= hi:
                                mid = (lo + hi) // 2
                                f_score = tkfont.Font(family='Arial', size=mid, weight='bold')
                                score_h = f_score.metrics('linespace')
                                # reserve proportions: name ~40%, score main ~50%, sub ~10%
                                name_h = max(10, int(score_h * 0.40))
                                sub_h = max(10, int(score_h * 0.22))
                                total_h = score_h + name_h + sub_h + 12
                                w_a = f_score.measure(score_text_a)
                                w_b = f_score.measure(score_text_b)
                                # allow each score up to ~45% of width (keep small gaps)
                                if (w_a <= w*0.45 and w_b <= w*0.45) and total_h <= h:
                                    best = mid
                                    lo = mid + 1
                                else:
                                    hi = mid - 1
                            f_score = tkfont.Font(family='Arial', size=best, weight='bold')
                            name_size = max(10, int(best * 0.5))
                            sub_size = max(10, int(best * 0.28))
                            tk.Label(frame, text=ten_a, font=('Arial', name_size, 'bold'), fg='white', bg='#000').pack(anchor='center', padx=6, pady=(6,0))
                            tk.Label(frame, text=f'TRẬN {tran}', font=('Arial', max(10, name_size-2), 'bold'), fg='#00E5FF', bg='#000').pack()
                            score_row = tk.Frame(frame, bg='#000')
                            score_row.pack(fill='both', expand=True)
                            tk.Label(score_row, text=score_text_a, font=('Arial', best, 'bold'), fg='white', bg='#000').pack(side='left', expand=True, padx=8)
                            tk.Label(score_row, text=score_text_b, font=('Arial', best, 'bold'), fg='#FFD600', bg='#000').pack(side='right', expand=True, padx=8)
                            sub = tk.Frame(frame, bg='#000')
                            sub.pack(fill='x', pady=6)
                            tk.Label(sub, text=f'AVG {avg_a}', fg='white', bg='#000', font=('Arial', sub_size)).pack(side='left', padx=6)
                            tk.Label(sub, text=f'Lco {lco}', fg='#FFD369', bg='#000', font=('Arial', sub_size)).pack(side='left', padx=6)
                            tk.Label(sub, text=f'AVG {avg_b}', fg='#FFD600', bg='#000', font=('Arial', sub_size)).pack(side='right', padx=6)
                        else:
                            tk.Label(frame, text='No vMix URL', fg='white', bg='#000').pack(padx=6, pady=6)
                    except Exception as ex:
                        tk.Label(frame, text=f'Error fetch vMix:\n{ex}', fg='red', bg='#000').pack(expand=True)
                else:
                    tk.Label(frame, text='Poster / Logo', fg='#FFD369', bg='#000', font=('Arial', 18, 'bold')).pack(expand=True)

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
                        render_cell_local(idx)
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
                    render_cell_local(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                def set_image():
                    path = filedialog.askopenfilename(filetypes=[('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp')])
                    if not path:
                        return
                    mode = simpledialog.askstring('Chế độ hiển thị', "Chọn 'center', 'fit' hoặc 'cover'", initialvalue='fit')
                    if mode not in ('center', 'fit', 'cover'):
                        mode = 'fit'
                    preview.cell_meta[idx] = {'type': 'image', 'value': path, 'image_ref': None, 'image_mode': mode}
                    render_cell_local(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                def clear_cell():
                    preview.cell_meta[idx] = {'type': None, 'value': None, 'image_ref': None}
                    render_cell_local(idx)
                    try:
                        self._update_last_preview_meta(preview)
                    except Exception:
                        pass

                tk.Button(dlg, text='Chọn từ bảng', command=set_vmix_from_row, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Nhập URL vMix', command=set_vmix_from_url, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Tải ảnh', command=set_image, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Xóa', command=clear_cell, width=30).pack(padx=8, pady=6)
                tk.Button(dlg, text='Đóng', command=dlg.destroy, width=30).pack(padx=8, pady=8)

            # Always create a fixed 3x3 grid, never resize or change number of cells
            for r in range(3):
                preview.grid_rowconfigure(r, weight=1, minsize=1)
                for c in range(3):
                    preview.grid_columnconfigure(c, weight=1, minsize=1)
                    i = r*3 + c
                    f = tk.Frame(preview, bg='#000', bd=4, relief='ridge')
                    f.grid(row=r, column=c, sticky='nsew', padx=4, pady=4)
                    preview.cells.append(f)
                    # Bind to <Configure> to auto-scale content inside cell
                    f.bind('<Configure>', lambda e, ii=i: render_cell_local(ii))
                    render_cell_local(i)

            ctrl = tk.Frame(preview, bg='#111')
            ctrl.grid(row=3, column=0, columnspan=3, sticky='ew')
            auto_var = tk.IntVar(value=1)
            try:
                tk.Checkbutton(ctrl, text='Auto-refresh', variable=auto_var, bg='#111', fg='white').pack(side='left', padx=8)
            except Exception:
                pass
            btn_close = tk.Button(ctrl, text='Đóng preview', command=lambda: preview.destroy(), bg='#FF6F00')
            btn_close.pack(side='right', padx=8, pady=6)

            def refresh_loop():
                try:
                    if auto_var.get():
                        for i, meta in enumerate(preview.cell_meta):
                            if meta and meta.get('type') == 'vmix':
                                try:
                                    render_cell_local(i)
                                except Exception:
                                    pass
                    preview.after(1500, refresh_loop)
                except Exception:
                    try:
                        preview.after(1500, refresh_loop)
                    except Exception:
                        pass

            # start refresh loop
            try:
                preview.after(1500, refresh_loop)
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

    def _auto_save_state(self):
        """Lưu trạng thái UI (sheet_url, credentials, ban, table rows, header fields) vào file pickle."""
        # Không autosave khi đang khôi phục trạng thái
        if getattr(self, '_restoring_state', False):
            return
        # If a restore is in-progress (not yet committed), skip autosave to avoid overwriting
        if not getattr(self, '_restore_committed', True):
            return
        try:
            state = {}
            try:
                state['tengiai'] = self.tengiai_var.get() if hasattr(self, 'tengiai_var') else ''
                state['thoigian'] = self.thoigian_var.get() if hasattr(self, 'thoigian_var') else ''
                state['diadiem'] = self.diadiem_var.get() if hasattr(self, 'diadiem_var') else ''
                state['chuchay'] = self.chuchay_var.get() if hasattr(self, 'chuchay_var') else ''
                state['diemso'] = self.diemso_text.get('1.0', 'end-1c') if hasattr(self, 'diemso_text') else ''
                state['sheet_url'] = self.url_var.get() if hasattr(self, 'url_var') else ''
                state['creds_path'] = self.creds_path if hasattr(self, 'creds_path') else None
                # Luôn lưu số bàn, kể cả khi là 0 hoặc None
                try:
                    state['ban'] = int(self.ban_var.get()) if hasattr(self, 'ban_var') and self.ban_var is not None else 0
                except Exception:
                    state['ban'] = 0
            except Exception:
                pass
            # serialize current table (first 6 columns) for each row
            table = []
            try:
                for r in getattr(self, 'match_rows', []):
                    rowvals = []
                    for j in range(6):
                        try:
                            w = r[j]
                            if hasattr(w, 'get'):
                                rowvals.append(w.get())
                            else:
                                rowvals.append('')
                        except Exception:
                            rowvals.append('')
                    table.append(rowvals)
            except Exception:
                table = []
            state['table'] = table
            # serialize preview configuration (use last-known meta if preview window closed)
            try:
                preview = getattr(self, '_preview_window', None)
                serial_preview = None
                if preview is not None and getattr(preview, 'winfo_exists', lambda: False)():
                    pm = getattr(preview, 'cell_meta', None)
                    if pm:
                        serial_preview = []
                        for cell in pm:
                            try:
                                serial_preview.append({
                                    'type': cell.get('type'),
                                    'value': cell.get('value'),
                                    'image_mode': cell.get('image_mode', 'fit')
                                })
                            except Exception:
                                serial_preview.append(None)
                        # update last-known snapshot
                        try:
                            self._last_preview_meta = serial_preview
                        except Exception:
                            pass
                else:
                    # use previously stored preview meta if available
                    serial_preview = getattr(self, '_last_preview_meta', None)
                try:
                    state['preview'] = serial_preview
                except Exception:
                    state['preview'] = None
            except Exception:
                try:
                    state['preview'] = getattr(self, '_last_preview_meta', None)
                except Exception:
                    state['preview'] = None
            # write to file
            try:
                # compare with existing file and skip write if identical
                should_write = True
                try:
                    if os.path.exists(self._auto_state_path):
                        import pickle as _pickle
                        try:
                            with open(self._auto_state_path, 'rb') as _f:
                                existing = _pickle.load(_f)
                            if existing == state:
                                should_write = False
                        except Exception:
                            # if loading fails, proceed to write
                            should_write = True
                except Exception:
                    should_write = True

                if not should_write:
                    try:
                        dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                        import datetime
                        with open(dbg, 'a', encoding='utf-8') as df:
                            df.write(f"[{datetime.datetime.now().isoformat()}] AUTO_SAVE SKIP -> {self._auto_state_path} keys={list(state.keys())} table_rows={len(state.get('table',[]))}\n")
                    except Exception:
                        pass
                else:
                    with open(self._auto_state_path, 'wb') as f:
                        import pickle
                        pickle.dump(state, f)
                    # log successful save
                    try:
                        dbg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vmix_debug.log')
                        import datetime
                        with open(dbg, 'a', encoding='utf-8') as df:
                            df.write(f"[{datetime.datetime.now().isoformat()}] AUTO_SAVE -> {self._auto_state_path} keys={list(state.keys())} table_rows={len(state.get('table',[]))}\n")
                    except Exception:
                        pass
            except Exception as ex:
                try:
                    print(f"[WARN] Không lưu được trạng thái UI: {ex}")
                except Exception:
                    pass
        except Exception:
            pass

    def _periodic_autosave(self):
        try:
            # if suspended due to recent restore, skip this save cycle
            try:
                import time
                if getattr(self, '_autosave_suspended_until', 0) and time.time() < self._autosave_suspended_until:
                    # reschedule quickly to recheck
                    try:
                        self.after(1000, self._periodic_autosave)
                    except Exception:
                        pass
                    return
            except Exception:
                pass
            self._auto_save_state()
        except Exception:
            pass
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
        import os, sys
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
                        'image_mode': cell.get('image_mode', 'fit')
                    })
                except Exception:
                    serial_preview.append(None)
            self._last_preview_meta = serial_preview
        except Exception:
            pass

    def save_preview_now(self):
        """Force-save the current preview configuration into the autosave file."""
        try:
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
        except Exception:
            pass

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
                        new_meta.append({'type': cell.get('type'), 'value': cell.get('value'), 'image_ref': None, 'image_mode': cell.get('image_mode', 'fit')})
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
        header_frame.pack(fill='x', pady=8)
        for i in range(0, 6):
            header_frame.grid_columnconfigure(i, weight=1)
        # Icon (if any)
        col_offset = 0
        if hasattr(self, 'icon_img'):
            tk.Label(header_frame, image=self.icon_img, bg='#222831').grid(row=0, column=0, rowspan=2, padx=8, pady=2, sticky='w')
            col_offset = 1
        # Row 0: TÊN GIẢI - THỜI GIAN - ĐIỂM SỐ
        tk.Label(header_frame, text='Tên giải:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=0, column=col_offset+0, sticky='e', padx=(10,4), pady=6)
        self.tengiai_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.tengiai_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369').grid(row=0, column=col_offset+1, sticky='ew', padx=4, pady=6)
        tk.Label(header_frame, text='Thời gian:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=0, column=col_offset+2, sticky='e', padx=4, pady=6)
        self.thoigian_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.thoigian_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369').grid(row=0, column=col_offset+3, sticky='ew', padx=4, pady=6)
        tk.Label(header_frame, text='Điểm số:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=0, column=col_offset+4, sticky='e', padx=4, pady=6)
        self.diemso_var = tk.StringVar()
        self.diemso_text = tk.Text(header_frame, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369', height=2, width=14, wrap='word')
        self.diemso_text.grid(row=0, column=col_offset+5, sticky='ew', padx=(4,10), pady=6)
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
        tk.Label(header_frame, text='Địa điểm:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=1, column=col_offset+0, sticky='e', padx=(10,4), pady=6)
        self.diadiem_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.diadiem_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369').grid(row=1, column=col_offset+1, columnspan=5, sticky='ew', padx=(4,10), pady=6)
        # Row 2: CHỮ CHẠY
        tk.Label(header_frame, text='Chữ chạy:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=2, column=col_offset+0, sticky='e', padx=(10,4), pady=6)
        self.chuchay_var = tk.StringVar()
        self.chuchay_text = tk.Text(header_frame, font=('Segoe UI', 16), relief='groove', bd=3, bg='#232B3E', fg='white', height=4, wrap='word', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369')
        self.chuchay_text.grid(row=2, column=col_offset+1, columnspan=5, sticky='ew', padx=(4,10), pady=6)
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

        # Sheet URL input frame
        url_frame = tk.Frame(self, bg='#232B3E')
        url_frame.pack(fill='x', pady=2)
        tk.Label(url_frame, text='Link Google Sheet:', fg='#FFD369', bg='#232B3E', font=('Segoe UI', 14, 'bold')).pack(side='left', padx=10)
        self.url_var = tk.StringVar()
        url_entry = tk.Entry(url_frame, textvariable=self.url_var, width=40, font=('Segoe UI', 14), relief='groove', bd=2, bg='#1A2233', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369')
        url_entry.pack(side='left', padx=5)
        url_entry.bind('<Return>', lambda e: self.reload_matches())
        # Thêm sự kiện double-click để mở link Google Sheet nếu có
        def open_gsheet_link(event=None):
            import webbrowser
            url = self.url_var.get().strip()
            if url:
                webbrowser.open(url)
        url_entry.bind('<Double-Button-1>', open_gsheet_link)
        tk.Button(url_frame, text='Chọn credentials', command=self.select_credentials, bg='#FFD369', fg='#1A2233', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#FFE082', activeforeground='#1A2233').pack(side='left', padx=5)
        self.creds_label = tk.Label(url_frame, text='(Chưa chọn credentials)', fg='#FFD369', bg='#232B3E', font=('Segoe UI', 11, 'italic'))
        self.creds_label.pack(side='left', padx=5)
        # Thêm Label hiển thị status_var
        self.status_var = tk.StringVar(value='')
        # ...existing code...
        tk.Button(url_frame, text='Tải dữ liệu', command=self.reload_matches, bg='#FFD369', fg='#1A2233', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#FFE082', activeforeground='#1A2233').pack(side='left', padx=10)
        # Nút preview dữ liệu Google Sheet
        # Đã có nút Kết quả ở từng bàn, không cần nút này ở dòng trên nữa
        def preview_sheet():
            import tkinter as tk
            import json
            win = tk.Toplevel(self)
            win.base_title = 'Preview dữ liệu Google Sheet'
            win.title(win.base_title)
            try:
                self.animate_title(win, win.base_title)
            except Exception:
                try:
                    win.title(win.base_title)
                except Exception:
                    pass
            win.geometry('800x400')
            text = tk.Text(win, font=('Consolas', 11), wrap='none')
            text.pack(fill='both', expand=True)
            if self.sheet_rows:
                try:
                    pretty = json.dumps(self.sheet_rows, ensure_ascii=False, indent=2)
                except Exception:
                    pretty = str(self.sheet_rows)
                text.insert('1.0', pretty)
            else:
                text.insert('1.0', 'Không có dữ liệu Google Sheet!')
            text.config(state='disabled')
        tk.Button(url_frame, text='Preview Sheet', command=preview_sheet, bg='#FFD369', fg='#1A2233', font=('Segoe UI', 11), relief='groove', bd=2, activebackground='#FFE082', activeforeground='#1A2233').pack(side='left', padx=5)
        tk.Button(url_frame, text='Làm mới', command=self.reload_matches, bg='#232B3E', fg='#FFD369', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#2C3650', activeforeground='#FFD369').pack(side='left', padx=5)
        tk.Button(url_frame, text='Force Write', command=lambda: self._force_write_prompt(), bg='#FF6F00', fg='white', font=('Segoe UI', 11, 'bold'), relief='groove', bd=2, activebackground='#FF8F00', activeforeground='white').pack(side='left', padx=5)
        # Place Save/Load buttons next to Force Write so they're visible and not overlapping header
        try:
            self.save_btn = tk.Button(url_frame, text='Lưu bảng', command=self.save_table_to_csv, bg='#00C853', fg='#1A2233', font=('Segoe UI', 11, 'bold'), relief='groove', bd=2, activebackground='#B9F6CA', activeforeground='#1A2233')
            self.load_btn = tk.Button(url_frame, text='Tải bảng', command=self.load_table_from_csv, bg='#2962FF', fg='#FFD369', font=('Segoe UI', 11, 'bold'), relief='groove', bd=2, activebackground='#82B1FF', activeforeground='#FFD369')
            self.save_btn.pack(side='left', padx=8)
            self.load_btn.pack(side='left', padx=4)
            # Add manual load-state button for debugging restore
            self.load_state_btn = tk.Button(url_frame, text='Load trạng thái', command=self.manual_load_state, bg='#FFAB00', fg='white', font=('Segoe UI', 11, 'bold'), relief='groove', bd=2)
            self.load_state_btn.pack(side='left', padx=6)
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
                # overlay controls
                ctrl = tk.Frame(frame, bg='#111')
                ctrl.place(relx=1.0, rely=0.0, anchor='ne')
                tk.Button(ctrl, text='⋮', width=2, command=lambda i=idx: config_cell(i), bg='#FFD369').pack()

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
                            tk.Label(frame, text='No vMix URL', fg='white', bg='#000').pack(padx=8, pady=8)
                    except Exception as ex:
                        tk.Label(frame, text=f'Error fetch vMix:\n{ex}', fg='red', bg='#000').pack(padx=8, pady=8)
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
                    # Ask display mode
                    mode = simpledialog.askstring('Chế độ hiển thị', "Chọn chế độ hiển thị: 'center', 'fit', hoặc 'cover' (mặc định: fit)")
                    if not mode or mode not in ('center', 'fit', 'cover'):
                        mode = 'fit'
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
        self.col_weights = [6, 4, 18, 18, 2, 9, 1, 6, 4]
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
        self.preview_all_btn = tk.Button(self.bottom_frame, text='Preview tổng hợp', command=open_preview_all, bg='#00B8D4', fg='#fff', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#4DD0E1', activeforeground='#222831')
        self.preview_all_btn.pack(side='right', padx=10)
        try:
            self.save_preview_btn = tk.Button(self.bottom_frame, text='Save preview now', command=self.save_preview_now, bg='#FF8A65', fg='white', font=('Segoe UI', 10, 'bold'), relief='groove', bd=1)
            self.save_preview_btn.pack(side='right', padx=6)
        except Exception:
            pass
        # Ensure button invokes class method if available
        try:
            self.preview_all_btn.config(command=lambda: self.open_preview_all())
        except Exception:
            pass
        # expose function to instance so tests can call programmatically
        try:
            self.open_preview_all = open_preview_all
        except Exception:
            pass
        # Thêm dòng báo trạng thái bên phải
        self.status_label = tk.Label(self.bottom_frame, textvariable=self.status_var, fg='#00FF00', bg='#222831', font=('Segoe UI', 12, 'italic'))
        self.status_label.pack(side='right', padx=10)
        # Nút lấy tất cả từ vMix (non-blocking)
        btn_fetch_all = tk.Button(self.bottom_frame, text='Lấy tất cả từ vMix', command=self.fetch_all_vmix_to_table, bg='#00C853', fg='#222831', font=('Arial', 11, 'bold'))
        btn_fetch_all.pack(side='right', padx=10)
        try:
            btn_fetch_all.config(command=lambda: self.fetch_all_vmix_to_table_async())
        except Exception:
            pass
        # Nút ghi tất cả lên Google Sheets (mở popup preview trước khi ghi)
        try:
            btn_write_all = tk.Button(self.bottom_frame, text='Ghi tất cả lên Sheets', bg='#2962FF', fg='white', font=('Arial', 11, 'bold'))
            btn_write_all.pack(side='right', padx=10)
            try:
                btn_write_all.config(command=lambda: self.show_preview_write_popup())
            except Exception:
                try:
                    btn_write_all.config(command=lambda: self.write_all_vmix_to_sheet_async())
                except Exception:
                    btn_write_all.config(command=lambda: self.write_all_vmix_to_sheet())
        except Exception:
            pass
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
                    import os, datetime
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
        self.match_rows = []
        self.row_states = []
        for i in range(self.ban_var.get()):
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
            # Số bàn (giảm width, căn giữa)
            e_ban = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), relief='groove', bd=2, insertbackground='white', justify='center')
            if i < len(old_rows):
                e_ban.insert(0, old_rows[i][1])
            # Increase displayed width for 'Bàn' (scaled up)
            try:
                e_ban.config(width=8)
            except Exception:
                pass
            e_ban.grid(row=i+1, column=1, padx=2, pady=2, ipadx=6, ipady=6, sticky='ew')
            # Hiệu ứng dấu nháy trắng khi focus
            def on_focus_in_ban(event, entry=e_ban):
                try:
                    entry.config(insertbackground='white')
                except:
                    pass
            e_ban.bind('<FocusIn>', on_focus_in_ban)
            widgets.append(e_ban)
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
            e_diem.grid(row=i+1, column=4, padx=2, pady=2, ipadx=7, ipady=6, sticky='ew')
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
            # Nút Kết quả cập nhật Google Sheet cho từng bàn
            def update_gsheet_for_row(idx=i):
                try:
                    import requests, xml.etree.ElementTree as ET, sys, os
                    row = self.match_rows[idx]
                    tran_val = row[0].get().strip()
                    vmix_url = row[5].get().strip()
                    sheet_url = self.url_var.get().strip()
                    if not tran_val or not vmix_url or not sheet_url:
                        messagebox.showerror('Lỗi', 'Thiếu thông tin Trận, Địa chỉ vMix hoặc Google Sheet!')
                        return
                    m = __import__('re').search(r"/spreadsheets/d/([\w-]+)", sheet_url)
                    spreadsheet_id = m.group(1) if m else None
                    creds_path = self.creds_path if self.creds_path else getattr(fetch_matches_from_sheet, '_creds_path', None)
                    if not (GSheetClient and spreadsheet_id and creds_path and os.path.exists(creds_path)):
                        messagebox.showerror('Lỗi', 'Chưa cấu hình đúng Google Sheets!')
                        return
                    gs = GSheetClient(spreadsheet_id, creds_path)
                    read_range = 'Kết quả!A1:Z2000'
                    rows = gs.read_table(read_range)
                    headers = list(rows[0].keys()) if rows else []
                    def find_col_key(keys, *candidates):
                        def normalize(s):
                            return s.replace(' ', '').replace('_', '').lower()
                        norm_keys = {normalize(k): k for k in keys}
                        for c in candidates:
                            nc = normalize(c)
                            for nk, orig in norm_keys.items():
                                if nk == nc:
                                    return orig
                        return None
                    tran_col = find_col_key(headers, 'Trận')
                    diem_a_col = find_col_key(headers, 'Điểm A', 'DiemA', 'DiemA.Text')
                    diem_b_col = find_col_key(headers, 'Điểm B', 'DiemB', 'DiemB.Text')
                    lco_col = find_col_key(headers, 'Lượt cơ', 'Lco', 'Lco.Text')
                    hr1a_col = find_col_key(headers, 'HR1A', 'HR1A.Text')
                    hr2a_col = find_col_key(headers, 'HR2A', 'HR2A.Text')
                    hr1b_col = find_col_key(headers, 'HR1B', 'HR1B.Text')
                    hr2b_col = find_col_key(headers, 'HR2B', 'HR2B.Text')
                    if not tran_col:
                        messagebox.showerror('Lỗi', 'Không tìm thấy cột Trận trong Google Sheet!')
                        return
                    input_tran_num = None
                    m = __import__('re').search(r'(\d+)', tran_val)
                    if m:
                        input_tran_num = int(m.group(1))
                    found_idx = None
                    for idx2, r in enumerate(rows):
                        sheet_tran_val = r.get(tran_col, '')
                        m2 = __import__('re').search(r'(\d+)', str(sheet_tran_val))
                        sheet_tran_num = int(m2.group(1)) if m2 else str(sheet_tran_val).strip().lower()
                        if input_tran_num == sheet_tran_num:
                            found_idx = idx2
                            break
                    if found_idx is None:
                        messagebox.showerror('Lỗi', 'Không tìm thấy dòng trận này trong Google Sheet!')
                        return
                    # Lấy dữ liệu vMix Input 1
                    resp = requests.get(f'{vmix_url}/API/', timeout=2)
                    xml = resp.text
                    root = ET.fromstring(xml)
                    input1 = root.find(".//input[@number='1']")
                    def get_field(name):
                        if input1 is not None:
                            for text in input1.findall('text'):
                                if text.attrib.get('name') == name:
                                    return text.text or ''
                        return ''
                    diem_a = get_field('DiemA.Text')
                    diem_b = get_field('DiemB.Text')
                    lco = get_field('Lco.Text')
                    hr1a = get_field('HR1A.Text')
                    hr2a = get_field('HR2A.Text')
                    hr1b = get_field('HR1B.Text')
                    hr2b = get_field('HR2B.Text')
                    update_row = rows[found_idx]
                    if diem_a_col:
                        update_row[diem_a_col] = diem_a
                    if diem_b_col:
                        update_row[diem_b_col] = diem_b
                    if lco_col:
                        update_row[lco_col] = lco
                    if hr1a_col:
                        update_row[hr1a_col] = hr1a
                    if hr2a_col:
                        update_row[hr2a_col] = hr2a
                    if hr1b_col:
                        update_row[hr1b_col] = hr1b
                    if hr2b_col:
                        update_row[hr2b_col] = hr2b
                    # Chỉ cập nhật các cột kết quả, giữ nguyên các cột khác
                    result_cols = set()
                    for col in [diem_a_col, diem_b_col, lco_col, hr1a_col, hr2a_col, hr1b_col, hr2b_col]:
                        if col:
                            result_cols.add(col)
                    new_values = [headers]
                    for idx2, r in enumerate(rows):
                        if idx2 == found_idx:
                            new_row = []
                            for h in headers:
                                if h in result_cols:
                                    new_row.append(update_row.get(h, ''))
                                else:
                                    new_row.append(r.get(h, ''))
                            new_values.append(new_row)
                        else:
                            new_values.append([r.get(h, '') for h in headers])
                    # Prepare contiguous AA..AK values and write only those columns for the matched row
                    # (TenA, TenB, Điểm A, Điểm B, Lượt cơ, HR1A, HR2A, HR1B, HR2B, AVGA, AVGB)
                    try:
                        # compute numeric match number for popup/skip logic
                        # proceed with writing for any match number

                        def idx_to_col(i):
                            n = i+1
                            s = ''
                            while n>0:
                                n, rem = divmod(n-1, 26)
                                s = chr(65+rem) + s
                            return s

                        # compute AVGA/AVGB as before
                        diem_a = update_row.get('Điểm A')
                        diem_b = update_row.get('Điểm B')
                        lco = update_row.get('Lượt cơ')
                        if diem_a is None:
                            diem_a = rows[found_idx].get('Điểm A')
                        if diem_b is None:
                            diem_b = rows[found_idx].get('Điểm B')
                        if lco is None:
                            lco = rows[found_idx].get('Lượt cơ')
                        def to_float(val):
                            try:
                                if val is None or val == '':
                                    return None
                                if isinstance(val, (int, float)):
                                    return float(val)
                                val = str(val).replace("'", "").replace(",", ".")
                                return float(val)
                            except Exception:
                                return None
                        a = to_float(diem_a)
                        b = to_float(diem_b)
                        c = to_float(lco)
                        avga = round(a/c, 3) if a is not None and c and c != 0 else ''
                        avgb = round(b/c, 3) if b is not None and c and c != 0 else ''

                        # Map vMix fields to specific sheet columns and write only those cells
                        ten_a_col = find_col_key(headers, 'Tên VĐV A', 'TenA', 'TenA.Text', 'VĐV A', 'VĐV_A')
                        ten_b_col = find_col_key(headers, 'Tên VĐV B', 'TenB', 'TenB.Text', 'VĐV B', 'VĐV_B')
                        avga_col = find_col_key(headers, 'AVGA', 'AvgA', 'AvgA.Text')
                        avgb_col = find_col_key(headers, 'AVGB', 'AvgB', 'AvgB.Text')

                        # values from vMix
                        ten_a_val = get_field('TenA') or get_field('TenA.Text') or update_row.get(ten_a_col, '')
                        ten_b_val = get_field('TenB') or get_field('TenB.Text') or update_row.get(ten_b_col, '')
                        vals_map = {
                            ten_a_col: ten_a_val,
                            ten_b_col: ten_b_val,
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

                        def index_to_col_letter(zero_based):
                            n = zero_based + 1
                            letters = ''
                            while n > 0:
                                n, rem = divmod(n-1, 26)
                                letters = chr(65 + rem) + letters
                            return letters

                        rownum = found_idx + 2
                        batch = []
                        for header_key, value in vals_map.items():
                            if not header_key:
                                continue
                            try:
                                col_idx = headers.index(header_key)
                            except ValueError:
                                continue
                            col_letter = index_to_col_letter(col_idx)
                            cell_range = f'Kết quả!{col_letter}{rownum}'
                            batch.append({'range': cell_range, 'values': [[value]]})

                        # log and execute
                        try:
                            import datetime, os
                            dbg_path = os.path.join(os.getcwd(), 'vmix_debug.log')
                            with open(dbg_path, 'a', encoding='utf-8') as df:
                                df.write(f"[{datetime.datetime.now()}] PREWRITE per-row tran={tran_val} target_row={rownum} items={len(batch)}\n")
                                for it in batch:
                                    df.write(f"[{datetime.datetime.now()}] PREWRITE_ITEM {it.get('range')} values={it.get('values')}\n")
                        except Exception:
                            pass

                        if batch:
                            try:
                                # prefer safe batch update that only writes AA..AL
                                res = gs.batch_update_safe(batch)
                                # batch_update_safe returns {} when nothing was allowed/filtered;
                                # in that case fall back to full batch_update so user-visible columns are updated
                                if not res:
                                    try:
                                        gs.batch_update(batch)
                                    except Exception:
                                        pass
                            except Exception:
                                # fallback to regular batch_update if safe method raises
                                try:
                                    gs.batch_update(batch)
                                except Exception:
                                    pass

                        try:
                            import datetime, os
                            dbg_path = os.path.join(os.getcwd(), 'vmix_debug.log')
                            with open(dbg_path, 'a', encoding='utf-8') as df:
                                df.write(f"[{datetime.datetime.now()}] WRITE_OK per-row tran={tran_val} target_row={rownum} wrote={len(batch)}\n")
                        except Exception:
                            pass

                        messagebox.showinfo('Thành công', f'Đã ghi kết quả trận {input_tran_num if input_tran_num is not None else tran_val} thành công')
                        self.status_var.set(f'Đã ghi kết quả trận {input_tran_num if input_tran_num is not None else tran_val} -> row {rownum}')
                        return
                    except Exception as ex:
                        messagebox.showerror('Lỗi', f'Lỗi khi ghi kết quả: {ex}')
                        return
                except Exception as ex:
                    messagebox.showerror('Lỗi', f'Lỗi cập nhật Google Sheet: {ex}')
            # --- Nút Kết quả ---
            btn_ketqua = tk.Button(self.table_frame, text='Kết quả', bg='#00C853', fg='white', font=('Arial', 18, 'bold'),
                                   relief='raised', bd=2, width=10)
            btn_ketqua.config(width=6)
            btn_ketqua.grid(row=i+1, column=6, padx=2, pady=2, ipadx=0)
            widgets.append(btn_ketqua)
            # --- Nút Gửi ---
            btn_gui = tk.Button(self.table_frame, text='Gửi', bg='#00C853', fg='white', font=('Arial', 18, 'bold'),
                               relief='raised', bd=2, width=10)
            btn_gui.config(width=6)
            btn_gui.grid(row=i+1, column=7, padx=2, pady=2, ipadx=0)
            widgets.append(btn_gui)
            # --- Nút Sửa ---
            btn_sua = tk.Button(self.table_frame, text='Sửa', bg='#FF9800', fg='#222831', font=('Arial', 18, 'bold'),
                               command=lambda idx=i: self.open_edit_popup(idx), relief='raised', bd=2, width=10)
            btn_sua.config(width=6)
            btn_sua.grid(row=i+1, column=8, padx=2, pady=2, ipadx=0)
            widgets.append(btn_sua)

            # --- Logic đổi màu nút ---
            def set_btn_color(btn, state):
                if state == 'pending':
                    btn.config(bg='#00C853', fg='white')  # Xanh lá
                elif state == 'edit':
                    btn.config(bg='#388E3C', fg='white')  # Xanh đậm khi sửa
                elif state == 'success':
                    btn.config(bg='#D32F2F', fg='white')  # Đỏ khi gửi thành công
                elif state == 'fail':
                    btn.config(bg='black', fg='white')    # Đen khi lỗi

            set_btn_color(btn_gui, 'pending')
            set_btn_color(btn_ketqua, 'pending')

            # --- Theo dõi thay đổi tên A/B để đổi màu nút ---
            def on_edit_name(event=None, btns=(btn_gui, btn_ketqua)):
                for b in btns:
                    set_btn_color(b, 'edit')
            widgets[2].bind('<KeyRelease>', on_edit_name)
            widgets[3].bind('<KeyRelease>', on_edit_name)

            # --- Nút Gửi: gửi xong đổi màu ---
            def on_push_to_vmix(idx=i, btn=btn_gui):
                try:
                    self.push_to_vmix(idx)
                    set_btn_color(btn, 'success')
                except Exception:
                    set_btn_color(btn, 'fail')
            btn_gui.config(command=lambda idx=i, btn=btn_gui: on_push_to_vmix(idx, btn))

            # --- Nút Kết quả: gửi xong đổi màu ---
            def on_btn_ketqua(idx=i, btn=btn_ketqua):
                try:
                    update_gsheet_for_row(idx)
                    set_btn_color(btn, 'success')
                except Exception:
                    set_btn_color(btn, 'fail')
            btn_ketqua.config(command=lambda idx=i, btn=btn_ketqua: on_btn_ketqua(idx, btn))
            # Không còn bôi màu khi click vào ô
            self.match_rows.append(widgets)
            self.row_states.append(row_state)
        self.selected_row = None

    def update_vdv_from_tran(self, row_idx):
        """Khi nhập số trận, tự động lấy tên VĐV từ sheet và cập nhật dòng giao diện."""
        import sys
        if not self.sheet_rows or not isinstance(self.sheet_rows, list) or not self.sheet_rows[0]:
            self.status_var.set('Chưa có dữ liệu Google Sheet!')
            print('DEBUG: sheet_rows is empty or invalid', file=sys.stderr)
            return
        widgets = self.match_rows[row_idx]
        tran_val = widgets[0].get().strip()
        ban_val = widgets[1].get().strip()
        # Kiểm tra trùng số trận hoặc số bàn
        is_tran_duplicate = False
        is_ban_duplicate = False
        for i, row in enumerate(self.match_rows):
            if i != row_idx:
                if row[0].get().strip() == tran_val and tran_val:
                    is_tran_duplicate = True
                if row[1].get().strip() == ban_val and ban_val:
                    is_ban_duplicate = True
        status_label = widgets[-1]
        if is_tran_duplicate:
            from tkinter import messagebox
            widgets[0].delete(0, 'end')
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            status_label.config(text='Số trận đã bị trùng!', fg='red')
            messagebox.showwarning('Cảnh báo', 'Số trận này đã bị trùng! Vui lòng nhập số khác.')
            widgets[0].focus_set()
            return
        if is_ban_duplicate:
            from tkinter import messagebox
            widgets[1].delete(0, 'end')
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            status_label.config(text='Số bàn đã bị trùng!', fg='red')
            messagebox.showwarning('Cảnh báo', 'Số bàn này đã bị trùng! Vui lòng nhập số khác.')
            widgets[1].focus_set()
            return
        # Nếu số trận bị trùng thì không lấy tên VĐV, xoá ô tên VĐV
        if hasattr(widgets[0], 'is_duplicate') and widgets[0].is_duplicate:
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            return
        # Nếu ô Trận bị xóa trắng hoặc bị xóa ký tự (delete/backspace) thì xóa luôn tên VĐV
        if not tran_val:
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            # Xóa luôn trạng thái dòng
            if len(widgets) > 8:
                widgets[-1].config(text='', fg='#005A9E')
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
        # Prefer fetching VĐV from sheet 'KET QUA'
        try:
            ketqua_rows = fetch_matches_from_ketqua(self.sheet_url if hasattr(self, 'sheet_url') else self.url_var.get(), self.num_ban if hasattr(self, 'num_ban') else None)
        except Exception:
            ketqua_rows = None
        if isinstance(ketqua_rows, dict) and 'error' in ketqua_rows:
            ketqua_rows = None
        source_rows = ketqua_rows if ketqua_rows else (self.sheet_rows if self.sheet_rows else [])
        if not source_rows:
            status_label.config(text='Không có dữ liệu trong KET QUA hoặc Kết quả!', fg='red')
            return
        keys = list(source_rows[0].keys())
        print(f'DEBUG: source_rows[0] keys = {keys}', file=sys.stderr)
        # Cho phép người dùng chỉnh tên cột tại đây (tìm trong sheet 'KET QUA'):
        tran_col = find_col_key(keys, 'Trận', 'Số trận', 'B', 'C')
        vdv_a_col = find_col_key(keys, 'VĐVA', 'vđv a', 'D', 'F')
        vdv_b_col = find_col_key(keys, 'VĐVB', 'vđv b', 'F', 'H')
        print(f'DEBUG: tran_col={tran_col}, vdv_a_col={vdv_a_col}, vdv_b_col={vdv_b_col}', file=sys.stderr)
        status_label = widgets[-1]
        if not tran_col or not vdv_a_col or not vdv_b_col:
            status_label.config(text='Không tìm thấy tiêu đề cột!', fg='red')
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
            print(f'DEBUG: tran_col={tran_col}, sheet_rows keys: {[list(r.keys()) for r in self.sheet_rows]}', file=sys.stderr)
        if not found:
            print('DEBUG: Không tìm thấy trận phù hợp', file=sys.stderr)
            # Nếu không tìm thấy thì xóa tên VĐV, báo lỗi rõ ràng
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].config(state='readonly', fg='#222831')
            status_label.config(text='Không tìm thấy trận này trong Google Sheet!', fg='red')
        else:
            vdv_a = found.get(vdv_a_col, '') if vdv_a_col else ''
            vdv_b = found.get(vdv_b_col, '') if vdv_b_col else ''
            # Luôn để màu chữ tên VĐV là đen
            widgets[2].config(state='normal', fg='#222831'); widgets[2].delete(0, 'end'); widgets[2].insert(0, vdv_a); widgets[2].config(state='readonly', fg='#222831')
            widgets[3].config(state='normal', fg='#222831'); widgets[3].delete(0, 'end'); widgets[3].insert(0, vdv_b); widgets[3].config(state='readonly', fg='#222831')
            status_label.config(text='', fg='#005A9E')

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

    def reload_matches(self):
        self.sheet_url = self.url_var.get().strip()
        self.num_ban = self.ban_var.get()
        # Gán credentials path cho fetch_matches_from_sheet để ưu tiên dùng credentials đã chọn
        fetch_matches_from_sheet._creds_path = self.creds_path if self.creds_path else None
        # Lấy dữ liệu Google Sheet nếu có link
        if self.sheet_url:
            sheet_rows = fetch_matches_from_sheet(self.sheet_url, self.num_ban)
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
                self.sheet_rows = sheet_rows
                self.status_var.set('Đã tải dữ liệu Google Sheet!')
            else:
                self.sheet_rows = []
                self.status_var.set('Không có dữ liệu Google Sheet!')
        else:
            self.sheet_rows = []
            self.status_var.set('Chưa nhập link Google Sheet!')
        # Chỉ xóa bảng và điền dữ liệu mới nếu thực sự có dữ liệu mới từ Google Sheet
        if self.sheet_rows:
            self.populate_table()
        else:
            self.status_var.set('Không có dữ liệu Google Sheet! (Bảng cũ được giữ nguyên)')

    def push_to_vmix(self, idx):
        import requests, sys
        widgets = self.match_rows[idx]
        entry_tran = widgets[0]
        status_label = widgets[-1]
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
        try:
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=TenA.Text&Value={vdv_a}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=TenB.Text&Value={vdv_b}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=Tran.Text&Value={tran_fmt}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=Noi dung.Text&Value={diem}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=backdrop.gtzip&SelectedName=tieu de.Text&Value={self.tengiai_var.get()}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=backdrop.gtzip&SelectedName=Ten A.Text&Value={vdv_a}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=backdrop.gtzip&SelectedName=Ten B.Text&Value={vdv_b}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=backdrop.gtzip&SelectedName=thoi gian.Text&Value={self.thoigian_var.get()}', timeout=2)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=backdrop.gtzip&SelectedName=dia diem.Text&Value={self.diadiem_var.get()}', timeout=2)
        except Exception as ex:
            print(f'ERROR: Không gửi được dữ liệu lên vMix: {ex}', file=sys.stderr)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=backdrop.gtzip&SelectedName=noi dung.Text&Value={diem}', timeout=2)
            # --- Gửi vào Input ket qua.gtzip ---
            # Tên giải vào tieu de.Text
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=ket qua.gtzip&SelectedName=tieu de.Text&Value={self.tengiai_var.get()}', timeout=2)
            # Không gửi tên VĐV A/B vào input ket qua.gtzip nữa
            # --- Gửi vào Input chay chu ---
            # Chạy chữ vào Input=chay chu (Ticker1.Text)
            requests.get(f'{vmix_url}/API/?Function=SetText&Input=chay chu&SelectedName=Ticker1.Text&Value={self.chuchay_var.get()}', timeout=2)
            status_label.config(text='Đã gửi!', fg='#00FF00')
            self.status_var.set(f'Đã gửi {tran_fmt} ({ban}) lên vMix: {vmix_url}')
        except Exception as e:
            status_label.config(text='Lỗi gửi!', fg='red')
            # Nếu tran_fmt rỗng, chỉ báo lỗi chung
            if tran_fmt:
                self.status_var.set(f'Lỗi gửi {tran_fmt}: {e}')
            else:
                self.status_var.set(f'Lỗi gửi: {e}')


    def save_table_to_csv(self):
        # Save all header info, including multiline score
        self.diemso_var.set(self.diemso_text.get('1.0', 'end-1c'))
        import csv
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')], title='Lưu bảng ra file CSV')
        if not file_path:
            return
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                import csv
                writer = csv.writer(f, quoting=csv.QUOTE_ALL)
                # Save global parameters as header rows
                writer.writerow(['# TenGiai', self.tengiai_var.get()])
                writer.writerow(['# ChuChay', self.chuchay_var.get()])
                writer.writerow(['# ThoiGian', self.thoigian_var.get()])
                writer.writerow(['# DiaDiem', self.diadiem_var.get()])
                writer.writerow(['# DiemSo', self.diemso_text.get('1.0', 'end-1c')])
                writer.writerow(['# LinkGoogleSheet', self.url_var.get().strip()])
                writer.writerow(['# CredentialsPath', self.creds_path if self.creds_path else ''])
                writer.writerow(['# TotalTables', self.ban_var.get()])
                writer.writerow(['# DateSaved', __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                writer.writerow(['Số trận', 'Số bàn', 'Tên VĐV A', 'Tên VĐV B', 'Điểm số', 'Địa chỉ vMix'])
                for row in self.match_rows:
                    # Đảm bảo luôn đúng 6 cột, không có cột thừa hoặc thiếu
                    values = [row[0].get(), row[1].get(), row[2].get(), row[3].get(), row[4].get(), row[5].get()]
                    values = values[:6] + [''] * (6 - len(values))
                    writer.writerow(values)
            self.status_var.set(f'Đã lưu bảng ra {file_path}')
        except Exception as e:
            self.status_var.set(f'Lỗi lưu CSV: {e}')

    def load_table_from_csv(self):
        import csv
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')], title='Tải bảng từ file CSV')
        if not file_path:
            return
        with open(file_path, 'r', encoding='utf-8') as f:
            import csv
            reader = csv.reader(f)
            rows = list(reader)
        # Parse global parameters from header rows
        header_idx = 0
        total_tables = None
        diemso_loaded = ''
        for i in range(min(15, len(rows))):
            if rows[i][0].startswith('# TenGiai'):
                self.tengiai_var.set(rows[i][1])
            elif rows[i][0].startswith('# ChuChay'):
                self.chuchay_var.set(rows[i][1])
            elif rows[i][0].startswith('# ThoiGian'):
                self.thoigian_var.set(rows[i][1])
            elif rows[i][0].startswith('# DiaDiem'):
                self.diadiem_var.set(rows[i][1])
            elif rows[i][0].startswith('# DiemSo'):
                diemso_loaded = rows[i][1]
            elif rows[i][0].startswith('# LinkGoogleSheet'):
                self.url_var.set(rows[i][1])
            elif rows[i][0].startswith('# CredentialsPath'):
                self.creds_path = rows[i][1]
                self.creds_label.config(text=self.creds_path if self.creds_path else '(Chưa chọn credentials)', fg='#00FF00' if self.creds_path else '#FFD369')
            elif rows[i][0].startswith('# TotalTables'):
                try:
                    total_tables = int(rows[i][1])
                    self.ban_var.set(total_tables)
                except:
                    total_tables = None
            header_idx = i
            # Find the table header row
            # Tìm dòng tiêu đề bảng: chấp nhận dòng đầu tiên có đúng 6 cột, cột đầu là 'Số trận' (không phân biệt hoa thường, loại bỏ BOM, khoảng trắng)
            table_start = None
            for i in range(len(rows)):
                if not rows[i]:
                    continue
                first_col = rows[i][0].strip().replace('\ufeff', '').lower()
                if first_col == 'số trận' and len(rows[i]) == 6:
                    table_start = i
                    break
            if table_start is None:
                from tkinter import messagebox
                # Hiển thị nội dung file CSV (20 dòng đầu) để người dùng kiểm tra trực tiếp
                preview = '\n'.join([','.join(row) for row in rows[:20]])
                messagebox.showerror('File CSV không hợp lệ', f'Không tìm thấy tiêu đề bảng hợp lệ trong file CSV này.\n\nNội dung file (20 dòng đầu):\n---\n{preview}\n---\n\nHãy kiểm tra lại định dạng, dấu phẩy, ký tự lạ hoặc gửi file này cho kỹ thuật viên để hỗ trợ.')
                return
            # Chỉ clear bảng một lần, sau đó populate lại đúng số dòng
            self.clear_table()
            if total_tables:
                self.ban_var.set(total_tables)
            self.populate_table()
            # Gán lại giá trị từng ô từ file CSV ngay sau khi populate_table, không gọi lại populate_table nữa
            if hasattr(self, 'diemso_text'):
                self.diemso_text.delete('1.0', 'end')
                self.diemso_text.insert('1.0', diemso_loaded)
            data_rows = rows[table_start+1:table_start+1+len(self.match_rows)]
            for i, row in enumerate(data_rows):
                if not any(cell.strip() for cell in row):
                    continue
                if len(row) < 6:
                    row = row + [''] * (6 - len(row))
                # Gán lại giá trị từng ô cho đúng dòng
                for j in range(6):
                    if i < len(self.match_rows) and j < len(self.match_rows[i]):
                        self.match_rows[i][j].config(state='normal')
                        self.match_rows[i][j].delete(0, 'end')
                        self.match_rows[i][j].insert(0, row[j])
                        if j in [2,3]:
                            self.match_rows[i][j].config(state='readonly')
                    # The following code block is commented out due to undefined variables (widgets, bg, contrast_fg, old)
                    # widgets.append(e_ban)
                    # # Tên VĐV A
                    # e_a = tk.Entry(self.table_frame, font=('Arial', 18, 'bold'), bg=bg, fg=contrast_fg(bg), relief='groove', bd=2)
                    # e_a.insert(0, old[2])
                    # e_a.config(state='readonly', fg=contrast_fg(bg))
                    # e_a.grid(row=i+1, column=2, padx=2, pady=2, ipadx=6, ipady=8, sticky='ew')
                    # widgets.append(e_a)
                    # # Tên VĐV B
                    # e_b = tk.Entry(self.table_frame, font=('Arial', 18, 'bold'), bg=bg, fg=contrast_fg(bg), relief='groove', bd=2)
                    # e_b.insert(0, old[3])
                    # e_b.config(state='readonly', fg=contrast_fg(bg))
                    # e_b.grid(row=i+1, column=3, padx=2, pady=2, ipadx=6, ipady=8, sticky='ew')
                    # widgets.append(e_b)
                    # # Điểm số
                    # e_diem = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), justify='center', relief='groove', bd=2)
                    # e_diem.insert(0, old[4])
                    # e_diem.grid(row=i+1, column=4, padx=2, pady=2, ipadx=6, ipady=8, sticky='ew')
                    # def diem_user_edit(event, entry=e_diem):
                    #     entry._user_edited = True
                    # e_diem.bind('<Key>', diem_user_edit)
                    # widgets.append(e_diem)
                    # # Địa chỉ vMix cho từng dòng (giữ lại giá trị cũ nếu có)
                    # e_vmix = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), relief='groove', bd=2)
                    # e_vmix.insert(0, old[5] if old[5] else 'http://127.0.0.1:8088')
                    # e_vmix.grid(row=i+1, column=5, padx=2, pady=2, ipadx=6, ipady=8, sticky='ew')
                    # widgets.append(e_vmix)

if __name__ == '__main__':
    print('DEBUG: Starting FullScreenMatchGUI...')
    app = FullScreenMatchGUI()
    print('DEBUG: Entering mainloop...')
    app.mainloop()
    print('DEBUG: mainloop exited.')
