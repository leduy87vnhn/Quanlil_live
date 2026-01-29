import sys, os
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import simpledialog
import threading
try:
    import yaml
except ImportError:
    yaml = None
from gsheet_client import GSheetClient

# --- Google Sheets fetch integration ---
def fetch_matches_from_sheet(sheet_url, num_ban):
    """
    Fetch match data from Google Sheets sheet 'Kết quả' using service account.
    Mapping: cột 'Trận' -> tran, 'Bàn' -> ban, 'VĐV A' -> vdv_a, 'VĐV B' -> vdv_b
    """
    import re
    print(f'DEBUG: fetch_matches_from_sheet called with sheet_url={sheet_url}, creds_path={getattr(fetch_matches_from_sheet, "_creds_path", None)}', file=sys.stdout)
    m = re.search(r"/spreadsheets/d/([\w-]+)", sheet_url)
    spreadsheet_id = m.group(1) if m else None
    print(f'DEBUG: spreadsheet_id={spreadsheet_id}', file=sys.stdout)
    # Always use sheet 'Kết quả', lấy đủ 2000 dòng
    read_range = 'Kết quả!A1:Z2000'
    creds_path = getattr(fetch_matches_from_sheet, '_creds_path', None)
    import shutil, tempfile
    use_temp_cred = False
    temp_cred_path = None
    # Luôn giải nén credentials.json ra file tạm khi chạy exe (dù có file rời hay không)
    if hasattr(sys, '_MEIPASS'):
        src_cred = os.path.join(sys._MEIPASS, 'credentials.json')
    else:
        # Tìm ở src trước, nếu không có thì tìm ở thư mục gốc dự án
        src_dir = os.path.dirname(os.path.abspath(__file__))
        root_dir = os.path.abspath(os.path.join(src_dir, '..'))
        src_cred = os.path.join(src_dir, 'credentials.json')
        if not os.path.exists(src_cred):
            src_cred = os.path.join(root_dir, 'credentials.json')
    if os.path.exists(src_cred):
        temp_dir = tempfile.gettempdir()
        temp_cred_path = os.path.join(temp_dir, 'temp_credentials.json')
        try:
            shutil.copy2(src_cred, temp_cred_path)
            creds_path = temp_cred_path
            use_temp_cred = True
        except Exception as ex:
            print(f'DEBUG: Error copying credentials.json to temp: {ex}', file=sys.stdout)
    print(f'DEBUG: creds_path={creds_path}', file=sys.stdout)
    # Nếu vẫn không có, thử config.yaml (dự phòng)
    if not creds_path or not os.path.exists(creds_path):
        config_path = os.path.join(os.path.dirname(__file__), '..', 'config.yaml')
        config = None
        if os.path.exists(config_path) and yaml:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
        if config:
            if not spreadsheet_id:
                spreadsheet_id = config.get('sheets', {}).get('spreadsheet_id')
            if not creds_path:
                creds_path = config.get('sheets', {}).get('credentials_path')
    if not GSheetClient:
        return {'error': 'Không import được GSheetClient. Kiểm tra lại gsheet_client.py!'}
    if not spreadsheet_id:
        return {'error': 'Không tìm thấy spreadsheet_id từ URL Google Sheet!'}
    if not creds_path or not os.path.exists(creds_path):
        return {'error': f'Không tìm thấy file credentials.json tại {creds_path}!'}
    try:
        import shutil, tempfile
        use_temp_cred = False
        temp_cred_path = None
        if hasattr(sys, '_MEIPASS') and creds_path.startswith(sys._MEIPASS):
            temp_dir = tempfile.gettempdir()
            temp_cred_path = os.path.join(temp_dir, 'temp_credentials.json')
            shutil.copy2(creds_path, temp_cred_path)
            use_temp_cred = True
        cred_to_use = temp_cred_path if use_temp_cred and temp_cred_path else creds_path
        gs = GSheetClient(spreadsheet_id, cred_to_use)
        rows = gs.read_table(read_range)
        if use_temp_cred and temp_cred_path:
            try:
                os.remove(temp_cred_path)
            except Exception:
                pass
        if rows:
            keys = list(rows[0].keys())
            def norm(s):
                import unicodedata
                s = s.strip().lower()
                s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
                s = s.replace(' ', '').replace('_', '')
                return s
            norm_keys = [norm(k) for k in keys]
            tran_candidates = ['tran', 'trận', 'tranda', 'tran dau', 'trandau', 'so tran', 'match']
            vdv_a_candidates = ['vdva', 'vđva', 'tena', 'vdv a', 'vđv a', 'playera', 'namea']
            vdv_b_candidates = ['vdvb', 'vđvb', 'tenb', 'vdv b', 'vđv b', 'playerb', 'nameb']
            found_tran = any(any(norm(k) == norm(c) for c in tran_candidates) for k in keys)
            found_vdva = any(any(norm(k) == norm(c) for c in vdv_a_candidates) for k in keys)
            found_vdvb = any(any(norm(k) == norm(c) for c in vdv_b_candidates) for k in keys)
            if found_tran or found_vdva or found_vdvb:
                return rows
            else:
                return {'error': 'Google Sheet không có cột Trận, VĐVA, VĐVB (linh hoạt)!'}
        else:
            return {'error': 'Không có dòng nào trong Google Sheet!'}
    except Exception as ex:
        return {'error': f'Lỗi khi truy cập Google Sheet: {ex}'}
    # Fallback: dummy data nếu không có dữ liệu Google Sheet
    print('DEBUG: Using dummy data', file=sys.stdout)
    matches = []
    for i in range(1, num_ban+1):
        matches.append({
            'ban': f'BÀN {i:02d}',
            'tran': f'TRẬN {i:02d}',
            'vdv_a': f'VĐV A {i}',
            'vdv_b': f'VĐV B {i}',
        })
    return matches

# export_vmix_input1_to_tempfile moved into FullScreenMatchGUI class

class FullScreenMatchGUI(tk.Tk):
    def export_vmix_input1_to_tempfile(self):
        """Export dữ liệu Input 1 của từng bàn ra file vmix_input1_temp.csv với đầy đủ trường chuẩn cho Google Sheet"""
        import csv, os, requests, xml.etree.ElementTree as ET, sys
        fieldnames = [
            'TenA', 'TenB', 'DiemA', 'DiemB', 'Lco',
            'HR1A', 'HR2A', 'HR1B', 'HR2B', 'AvgA', 'AvgB'
        ]
        rows = []
        for row in self.match_rows:
            vmix_url = ''
            try:
                vmix_url = row[5].get().strip() if hasattr(row[5], 'get') else ''
            except Exception:
                vmix_url = ''
            if not vmix_url:
                rows.append({fn: '' for fn in fieldnames})
                continue
            try:
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
                rowdata = {
                    'TenA': get_field('TenA.Text'),
                    'TenB': get_field('TenB.Text'),
                    'DiemA': get_field('DiemA.Text'),
                    'DiemB': get_field('DiemB.Text'),
                    'Lco': get_field('Lco.Text'),
                    'HR1A': get_field('HR1A.Text'),
                    'HR2A': get_field('HR2A.Text'),
                    'HR1B': get_field('HR1B.Text'),
                    'HR2B': get_field('HR2B.Text'),
                    'AvgA': get_field('AvgA.Text'),
                    'AvgB': get_field('AvgB.Text'),
                }
                rows.append(rowdata)
            except Exception as ex:
                rows.append({fn: '' for fn in fieldnames})
        if hasattr(sys, '_MEIPASS'):
            exe_dir = os.path.dirname(sys.executable)
        else:
            exe_dir = os.path.dirname(os.path.abspath(__file__))
        temp_path = os.path.join(exe_dir, 'vmix_input1_temp.csv')
        try:
            with open(temp_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for r in rows:
                    writer.writerow(r)
            # Also write copies to current working directory and system temp for easier discovery
            try:
                import tempfile
                alt_path = os.path.join(os.getcwd(), 'vmix_input1_temp.csv')
                with open(alt_path, 'w', encoding='utf-8', newline='') as f2:
                    writer2 = csv.DictWriter(f2, fieldnames=fieldnames)
                    writer2.writeheader()
                    for r in rows:
                        writer2.writerow(r)
                tmp_alt = os.path.join(tempfile.gettempdir(), 'vmix_input1_temp.csv')
                with open(tmp_alt, 'w', encoding='utf-8', newline='') as f3:
                    writer3 = csv.DictWriter(f3, fieldnames=fieldnames)
                    writer3.writeheader()
                    for r in rows:
                        writer3.writerow(r)
            except Exception as e2:
                print(f'WARNING writing alternate vmix_input1_temp copies: {e2}', file=sys.stderr)
            print(f'DEBUG: vmix_input1_temp.csv written to: {temp_path}', file=sys.stdout)
            # Show a confirmation popup to the user with paths written
            try:
                import tkinter as tk
                from tkinter import messagebox
                msg = f"vmix_input1_temp.csv written to:\n{temp_path}\n{alt_path}\n{tmp_alt}"
                try:
                    messagebox.showinfo('vmix CSV created', msg, parent=self)
                except Exception:
                    # fallback: create temporary root
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showinfo('vmix CSV created', msg)
                    root.destroy()
            except Exception as e3:
                print(f'WARNING showing popup for vmix_input1_temp.csv: {e3}', file=sys.stderr)
        except Exception as e:
            print(f'ERROR writing vmix_input1_temp.csv: {e}', file=sys.stderr)
    def open_edit_popup(self, idx):
        """Mở popup sửa kết quả từng bàn, lấy dữ liệu từ vMix Input 1, cho phép sửa và gửi ngược lại vMix"""
        import tkinter as tk
        from tkinter import messagebox
        import requests
        import xml.etree.ElementTree as ET
        row = self.match_rows[idx]
        popup = tk.Toplevel(self)
        popup.title(f'Sửa kết quả bàn {idx+1}')
        popup.geometry('650x500')
        popup.grab_set()
        popup.transient(self)
        popup.resizable(False, False)
        # --- DARK THEME ---
        dark_bg = '#222831'
        entry_bg = '#393e46'
        fg = '#fff'
        label_fg = '#fff'
        popup.configure(bg=dark_bg)
        # Lấy địa chỉ vMix
        vmix_url = row[5].get().strip()
        # Lấy dữ liệu từ vMix Input 1
        diem_a = diem_b = lco = hr1a = hr2a = hr1b = hr2b = ten_a = ten_b = adi = bdi = ''
        error_msg = None
        try:
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
            ten_a = get_field('TenA.Text')
            ten_b = get_field('TenB.Text')
            diem_a = get_field('DiemA.Text')
            diem_b = get_field('DiemB.Text')
            lco = get_field('Lco.Text')
            hr1a = get_field('HR1A.Text')
            hr2a = get_field('HR2A.Text')
            hr1b = get_field('HR1B.Text')
            hr2b = get_field('HR2B.Text')
            adi = get_field('Adi.Text')
            bdi = get_field('Bdi.Text')
        except Exception as ex:
            error_msg = f'Lỗi lấy dữ liệu vMix: {ex}'
        # Layout: 2 cột, các trường kết quả
        font_label = ('Arial', 15, 'bold')
        font_entry = ('Arial', 15)
        def dark_label(*args, **kwargs):
            kwargs.setdefault('bg', dark_bg)
            kwargs.setdefault('fg', label_fg)
            return tk.Label(*args, **kwargs)
        def dark_entry(*args, **kwargs):
            kwargs.setdefault('bg', entry_bg)
            kwargs.setdefault('fg', fg)
            kwargs.setdefault('insertbackground', fg)
            return tk.Entry(*args, **kwargs)

        dark_label(popup, text='Tên A:', font=font_label).grid(row=0, column=0, sticky='e', padx=10, pady=8)
        e_ten_a = dark_entry(popup, font=font_entry, width=16)
        e_ten_a.insert(0, ten_a)
        e_ten_a.grid(row=0, column=1, padx=4, pady=8)
        dark_label(popup, text='Tên B:', font=font_label).grid(row=0, column=2, sticky='e', padx=10, pady=8)
        e_ten_b = dark_entry(popup, font=font_entry, width=16)
        e_ten_b.insert(0, ten_b)
        e_ten_b.grid(row=0, column=3, padx=4, pady=8)
        dark_label(popup, text='Điểm A:', font=font_label).grid(row=1, column=0, sticky='e', padx=10, pady=8)
        e_diem_a = dark_entry(popup, font=font_entry, width=8)
        e_diem_a.insert(0, diem_a)
        e_diem_a.grid(row=1, column=1, padx=4, pady=8)
        dark_label(popup, text='Điểm B:', font=font_label).grid(row=1, column=2, sticky='e', padx=10, pady=8)
        e_diem_b = dark_entry(popup, font=font_entry, width=8)
        e_diem_b.insert(0, diem_b)
        e_diem_b.grid(row=1, column=3, padx=4, pady=8)
        dark_label(popup, text='Lượt cơ:', font=font_label).grid(row=2, column=0, sticky='e', padx=10, pady=8)
        e_lco = dark_entry(popup, font=font_entry, width=8)
        e_lco.insert(0, lco)
        e_lco.grid(row=2, column=1, padx=4, pady=8)
        # Adi, Bdi ngay dưới dòng Lượt cơ (dòng 3)
        dark_label(popup, text='Adi:', font=font_label).grid(row=3, column=0, sticky='e', padx=10, pady=8)
        e_adi = dark_entry(popup, font=font_entry, width=8)
        e_adi.insert(0, adi)
        e_adi.grid(row=3, column=1, padx=4, pady=8)
        dark_label(popup, text='Bdi:', font=font_label).grid(row=3, column=2, sticky='e', padx=10, pady=8)
        e_bdi = dark_entry(popup, font=font_entry, width=8)
        e_bdi.insert(0, bdi)
        e_bdi.grid(row=3, column=3, padx=4, pady=8)
        # HR1A, HR1B
        dark_label(popup, text='HR1A:', font=font_label).grid(row=4, column=0, sticky='e', padx=10, pady=8)
        e_hr1a = dark_entry(popup, font=font_entry, width=8)
        e_hr1a.insert(0, hr1a)
        e_hr1a.grid(row=4, column=1, padx=4, pady=8)
        dark_label(popup, text='HR1B:', font=font_label).grid(row=4, column=2, sticky='e', padx=10, pady=8)
        e_hr1b = dark_entry(popup, font=font_entry, width=8)
        e_hr1b.insert(0, hr1b)
        e_hr1b.grid(row=4, column=3, padx=4, pady=8)
        # HR2A, HR2B
        dark_label(popup, text='HR2A:', font=font_label).grid(row=5, column=0, sticky='e', padx=10, pady=8)
        e_hr2a = dark_entry(popup, font=font_entry, width=8)
        e_hr2a.insert(0, hr2a)
        e_hr2a.grid(row=5, column=1, padx=4, pady=8)
        dark_label(popup, text='HR2B:', font=font_label).grid(row=5, column=2, sticky='e', padx=10, pady=8)
        e_hr2b = dark_entry(popup, font=font_entry, width=8)
        e_hr2b.insert(0, hr2b)
        e_hr2b.grid(row=5, column=3, padx=4, pady=8)
        if error_msg:
            tk.Label(popup, text=error_msg, fg='red', bg=dark_bg, font=('Arial', 12, 'italic')).grid(row=6, column=0, columnspan=4, pady=8)
        # Nút Gửi
        def send_to_vmix():
            try:
                # Gửi các trường đã sửa lên vMix Input 1
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=TenA.Text&Value={e_ten_a.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=TenB.Text&Value={e_ten_b.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=DiemA.Text&Value={e_diem_a.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=DiemB.Text&Value={e_diem_b.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=Lco.Text&Value={e_lco.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=Adi.Text&Value={e_adi.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=Bdi.Text&Value={e_bdi.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=HR1A.Text&Value={e_hr1a.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=HR2A.Text&Value={e_hr2a.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=HR1B.Text&Value={e_hr1b.get().strip()}', timeout=2)
                requests.get(f'{vmix_url}/API/?Function=SetText&Input=1&SelectedName=HR2B.Text&Value={e_hr2b.get().strip()}', timeout=2)
                messagebox.showinfo('Thành công', 'Đã cập nhật kết quả lên vMix!')
                popup.destroy()
            except Exception as ex:
                messagebox.showerror('Lỗi', f'Lỗi gửi dữ liệu lên vMix: {ex}')
        btn_send = tk.Button(popup, text='Gửi', font=('Arial', 15, 'bold'), bg='#00C853', fg='white', width=12, command=send_to_vmix, activebackground='#009624', activeforeground='white')
        btn_send.grid(row=7, column=0, columnspan=4, pady=18)
        popup.bind('<Return>', lambda e: send_to_vmix())
        popup.focus_set()
        popup.focus_set()
    # --- LẤY KẾT QUẢ VÀ CẬP NHẬT GOOGLE SHEET ---
    def update_gsheet_with_vmix_full(self):
        """Lấy DiemA, DiemB, Lco, HR1A, HR2A, HR1B, HR2B từ vMix từng bàn và cập nhật vào Google Sheet"""
        import requests
        import xml.etree.ElementTree as ET
        import sys
        import os
        sheet_url = self.url_var.get().strip()
        if not sheet_url or not self.sheet_rows:
            self.status_var.set('Chưa có dữ liệu Google Sheet!')
            return
        m = __import__('re').search(r"/spreadsheets/d/([\w-]+)", sheet_url)
        spreadsheet_id = m.group(1) if m else None
        creds_path = self.creds_path if self.creds_path else getattr(fetch_matches_from_sheet, '_creds_path', None)
        if not (GSheetClient and spreadsheet_id and creds_path and os.path.exists(creds_path)):
            self.status_var.set('Chưa cấu hình đúng Google Sheets!')
            return
        # Early debug marker so we always have a trace that the update was invoked
        try:
            import datetime
            dbg_path = os.path.join(os.getcwd(), 'vmix_debug.log')
            creds_exists = bool(creds_path and os.path.exists(creds_path))
            with open(dbg_path, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now()}] ENTER update_gsheet_with_vmix_full spreadsheet_id={spreadsheet_id} creds_exists={creds_exists}\n")
        except Exception:
            pass
        gs = GSheetClient(spreadsheet_id, creds_path)
        read_range = 'Kết quả!A1:Z2000'
        rows = gs.read_table(read_range)
        # Debug: log/read rows count to help diagnose why writes may be skipped
        try:
            import datetime
            dbg_path = os.path.join(os.getcwd(), 'vmix_debug.log')
            with open(dbg_path, 'a', encoding='utf-8') as df:
                df.write(f"[{datetime.datetime.now()}] update_gsheet_with_vmix_full: read_range={read_range} rows_count={len(rows) if rows else 0}\n")
        except Exception:
            pass
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
            self.status_var.set('Không tìm thấy cột Trận trong Google Sheet!')
            return
        # Tự động export file tạm trước khi đọc
        self.export_vmix_input1_to_tempfile()
        import csv, os
        if hasattr(sys, '_MEIPASS'):
            exe_dir = os.path.dirname(sys.executable)
        else:
            exe_dir = os.path.dirname(os.path.abspath(__file__))
        temp_path = os.path.join(exe_dir, 'vmix_input1_temp.csv')
        vmix_rows = []
        if os.path.exists(temp_path):
            with open(temp_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for r in reader:
                    vmix_rows.append(r)
        # Xác định các cột cần ghi
        # Luôn lấy cột D (index 3) cho VĐV A, cột F (index 5) cho VĐV B, nếu thiếu thì báo lỗi
        if len(headers) <= 5:
            self.status_var.set('Google Sheet phải có ít nhất 6 cột (D, F) để ghi VĐV A/B!')
            import tkinter as tk
            from tkinter import messagebox
            messagebox.showerror('Lỗi', 'Google Sheet phải có ít nhất 6 cột (D, F) để ghi VĐV A/B!')
            return
        vdv_a_col = headers[3]
        vdv_b_col = headers[5]
        avga_col = find_col_key(headers, 'AVGA', 'AvgA', 'AvgA.Text')
        avgb_col = find_col_key(headers, 'AVGB', 'AvgB', 'AvgB.Text')
        updated_cells = 0
        msg_lines = []
        for idx, sheet_row in enumerate(rows):
            try:
                if idx >= len(vmix_rows):
                    continue
                r = vmix_rows[idx]
                update = {}
                if diem_a_col: update[diem_a_col] = r.get('DiemA', '')
                if diem_b_col: update[diem_b_col] = r.get('DiemB', '')
                if lco_col: update[lco_col] = r.get('Lco', '')
                if hr1a_col: update[hr1a_col] = r.get('HR1A', '')
                if hr2a_col: update[hr2a_col] = r.get('HR2A', '')
                if hr1b_col: update[hr1b_col] = r.get('HR1B', '')
                if hr2b_col: update[hr2b_col] = r.get('HR2B', '')
                # Luôn ghi TenA vào cột D, TenB vào cột F
                update[vdv_a_col] = r.get('TenA', '')
                update[vdv_b_col] = r.get('TenB', '')
                if avga_col: update[avga_col] = r.get('AvgA', r.get('AVGA', r.get('AvgA.Text', '')))
                if avgb_col: update[avgb_col] = r.get('AvgB', r.get('AVGB', r.get('AvgB.Text', '')))
                # Ghi từng trường vào sheet nếu khác giá trị cũ
                for col_name, value in update.items():
                    if not col_name:
                        continue
                    old_value = sheet_row.get(col_name, '')
                    if str(value).strip() == str(old_value).strip():
                        continue
                    col_idx = headers.index(col_name)
                    col_letter = chr(65 + col_idx)
                    cell_range = f'Kết quả!{col_letter}{idx+2}'
                    gs.write_table(cell_range, [[value]])
                    updated_cells += 1
                # --- GHI BỔ SUNG: lưu TenA, TenB, DiemA, DiemB, Lco, HR1A, HR2A, HR1B, HR2B, AVGA, AVGB
                def index_to_col_letter(zero_based):
                    n = zero_based + 1
                    letters = ''
                    while n > 0:
                        n, rem = divmod(n-1, 26)
                        letters = chr(65 + rem) + letters
                    return letters
                extra_fields = [
                    ('TenA', r.get('TenA', '')),
                    ('TenB', r.get('TenB', '')),
                    ('DiemA', r.get('DiemA', '')),
                    ('DiemB', r.get('DiemB', '')),
                    ('Lco', r.get('Lco', '')),
                    ('HR1A', r.get('HR1A', '')),
                    ('HR2A', r.get('HR2A', '')),
                    ('HR1B', r.get('HR1B', '')),
                    ('HR2B', r.get('HR2B', '')),
                    ('AVGA', r.get('AvgA', r.get('AVGA', r.get('AvgA.Text', '')))),
                    ('AVGB', r.get('AvgB', r.get('AVGB', r.get('AvgB.Text', '')))),
                ]
                # Prepare batch writes for AA..AK only for empty cells to avoid overwriting formulas.
                start_idx = 26  # 0-based index for column AA
                # Determine max rows to probe for existing AA:AK values
                max_probe_rows = max(len(rows), len(vmix_rows)) + 1
                try:
                    probe_range = f'Kết quả!{index_to_col_letter(start_idx)}2:{index_to_col_letter(start_idx+10)}{max_probe_rows}'
                    existing_block = gs.batch_get(probe_range)
                except Exception:
                    existing_block = []

                batch_data = []
                for i, (fname, fval) in enumerate(extra_fields):
                    col_idx = start_idx + i
                    col_letter = index_to_col_letter(col_idx)
                    row_num = idx + 2
                    # existing_block is 0-based rows starting at row 2
                    existing_val = ''
                    try:
                        rpos = (row_num - 2)
                        if rpos < len(existing_block):
                            rowvals = existing_block[rpos]
                            # column within AA..AK block
                            if i < len(rowvals):
                                existing_val = rowvals[i]
                    except Exception:
                        existing_val = ''
                    cell_range = f'Kết quả!{col_letter}{row_num}'
                    try:
                        if str(existing_val).strip() == '':
                            # schedule this single-cell write in batch
                            batch_data.append({
                                'range': cell_range,
                                'values': [[fval]]
                            })
                            try:
                                import datetime
                                with open(os.path.join(os.getcwd(), 'vmix_write.log'), 'a', encoding='utf-8') as lf:
                                    lf.write(f"[{datetime.datetime.now()}] BATCH_SCHEDULE WRITE {cell_range} => {fval}\n")
                            except Exception:
                                pass
                            print(f"[DEBUG] SCHEDULE WRITE {cell_range} => {fval}")
                        else:
                            try:
                                import datetime
                                with open(os.path.join(os.getcwd(), 'vmix_write.log'), 'a', encoding='utf-8') as lf:
                                    lf.write(f"[{datetime.datetime.now()}] SKIP {cell_range} existing={existing_val}\n")
                            except Exception:
                                pass
                            print(f"[DEBUG] SKIP {cell_range} existing={existing_val}")
                    except Exception as ex:
                        print(f'ERROR preparing extra field {fname} row {idx}: {ex}', file=sys.stderr)

                # Execute batch update if any
                if batch_data:
                    try:
                        gs.batch_update(batch_data)
                        updated_cells += len(batch_data)
                        try:
                            import datetime
                            with open(os.path.join(os.getcwd(), 'vmix_write.log'), 'a', encoding='utf-8') as lf:
                                lf.write(f"[{datetime.datetime.now()}] BATCH_UPDATE sent {len(batch_data)} items\n")
                        except Exception:
                            pass
                    except Exception as ex:
                        print(f'ERROR executing batch update row {idx}: {ex}', file=sys.stderr)
                # Dù có cập nhật hay không, luôn thêm vào popup
                msg = f"Dòng {idx+2}: "
                msg += f"VĐV A: {r.get('TenA','')}, VĐV B: {r.get('TenB','')}, "
                msg += f"Điểm A: {r.get('DiemA','')}, Điểm B: {r.get('DiemB','')}, Lượt cơ: {r.get('Lco','')}"
                msg += f", HR1A: {r.get('HR1A','')}, HR2A: {r.get('HR2A','')}, HR1B: {r.get('HR1B','')}, HR2B: {r.get('HR2B','')}"
                msg += f", AVGA: {r.get('AvgA','')}, AVGB: {r.get('AvgB','')}"
                msg_lines.append(msg)
            except Exception as ex:
                print(f'ERROR update_gsheet_with_vmix_full row {idx}: {ex}', file=sys.stderr)
        self.status_var.set(f'Đã cập nhật {updated_cells} ô (Kết quả) lên Google Sheet!')
        # Luôn hiện popup thông tin vừa ghi
        import tkinter as tk
        from tkinter import messagebox
        msg_full = '\n'.join(msg_lines)
        messagebox.showinfo('Đã ghi vào Google Sheet', msg_full)

    def fetch_all_vmix_to_table(self):
        pass

        with ThreadPoolExecutor(max_workers=8) as executor:
            executor.map(fetch_one, self.match_rows)

    def _update_row_from_vmix(self, row, tran, ten_a, ten_b, diem, ex):
        if ex is None:
            row[0].delete(0, 'end')
            row[0].insert(0, tran)
            row[2].config(state='normal')
            row[2].delete(0, 'end')
            row[2].insert(0, ten_a)
            row[2].config(state='readonly')
            row[3].config(state='normal')
            row[3].delete(0, 'end')
            row[3].insert(0, ten_b)
            row[3].config(state='readonly')
            row[4].delete(0, 'end')
            row[4].insert(0, diem)
            row[-1].config(text='Đã lấy từ vMix', fg='#00C853')
        else:
            row[-1].config(text=f'Lỗi lấy vMix: {ex}', fg='red')

    def animate_title(self, window, base_title, colors=None, idx=0):
        # 7 emoji màu cầu vồng đại diện cho hiệu ứng chớp nháy
        rainbow_emojis = ['🟥', '🟧', '🟨', '🟩', '🟦', '🟪', '⬛']
        if colors is None:
            colors = rainbow_emojis
        # Tạo tiêu đề với emoji màu thay đổi liên tục
        color_emoji = colors[idx % len(colors)]
        rainbow_title = f"{color_emoji} {base_title} {color_emoji}"
        try:
            window.title(rainbow_title)
        except Exception:
            pass
        window.after(180, lambda: self.animate_title(window, base_title, colors, idx+1))

    def update_gsheet_with_vmix(self):
        """
        Lấy dữ liệu từ vMix Input 1 (Điểm, Lượt cơ, HR1, HR2) cho từng bàn và cập nhật vào đúng dòng/cột của sheet 'Kết quả' trên Google Sheets.
        """
        import requests
        import xml.etree.ElementTree as ET
        import sys
        sheet_url = self.url_var.get().strip()
        if not sheet_url:
            self.status_var.set('Bạn chưa nhập URL Google Sheet!')
            return
        # Lấy dữ liệu Google Sheet
        rows = fetch_matches_from_sheet(sheet_url, len(self.match_rows))
        error_msg = None
        if isinstance(rows, dict) and 'error' in rows:
            error_msg = rows['error']
        elif not rows:
            error_msg = 'Không có dữ liệu Google Sheet!\nHãy kiểm tra: URL, credentials.json, sheet phải có cột Trận, VĐVA, VĐVB, và đã cài đủ requirements.txt.'
        if error_msg:
            self.status_var.set(error_msg)
            try:
                with open('error_gsheet.log', 'a', encoding='utf-8') as f:
                    import datetime
                    f.write(f"[{datetime.datetime.now()}] {error_msg}\n")
            except Exception:
                pass
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = None
                if not hasattr(self, 'winfo_exists') or not self.winfo_exists():
                    root = tk.Tk()
                    root.withdraw()
                messagebox.showerror('Google Sheet Error', error_msg)
                if root:
                    root.destroy()
            except Exception:
                pass
            # In lỗi ra console để dễ kiểm tra khi chạy .exe hoặc python
            try:
                print(f"Google Sheet Error: {error_msg}", file=sys.stderr)
            except Exception:
                pass
            return
        # ...existing code...
            return
        m = __import__('re').search(r"/spreadsheets/d/([\w-]+)", sheet_url)
        spreadsheet_id = m.group(1) if m else None
        creds_path = self.creds_path if self.creds_path else getattr(fetch_matches_from_sheet, '_creds_path', None)
        if not (GSheetClient and spreadsheet_id and creds_path and os.path.exists(creds_path)):
            self.status_var.set('Chưa cấu hình đúng Google Sheets!')
            return
        gs = GSheetClient(spreadsheet_id, creds_path)
        # Xác định các cột cần cập nhật
        keys = list(self.sheet_rows[0].keys())
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
        tran_col = find_col_key(keys, 'Trận')
        diem_col = find_col_key(keys, 'Điểm', 'Điểm số', 'Noi dung', 'Noi dung.Text')
        luotco_col = find_col_key(keys, 'Lượt cơ', 'LuotCo', 'LuotCo.Text')
        hr1_col = find_col_key(keys, 'HR1', 'HR1.Text')
        hr2_col = find_col_key(keys, 'HR2', 'HR2.Text')
        if not tran_col:
            self.status_var.set('Không tìm thấy cột Trận trong Google Sheet!')
            return
        # Đọc lại toàn bộ sheet để lấy vị trí dòng
        read_range = 'Kết quả!A1:Z2000'
        rows = gs.read_table(read_range)
        headers = list(rows[0].keys()) if rows else []
        # Chuẩn bị cập nhật từng dòng
        updates = []
        vmix_map = {}
        for i, row in enumerate(self.match_rows):
            try:
                tran_val = row[0].get().strip()
                vmix_url = row[5].get().strip()
                if not tran_val or not vmix_url:
                    continue
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
                    continue
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
                import unicodedata
                def norm_name(s):
                    s = s.strip().lower()
                    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
                    s = s.replace('(', '').replace(')', '').replace('.', '').replace('-', ' ')
                    s = ' '.join(s.split())
                    return s
                def last_word(s):
                    return norm_name(s).split()[-1] if s else ''
                ten_a_vmix = get_field('TenA.Text')
                ten_b_vmix = get_field('TenB.Text')
                diem_a_vmix = get_field('DiemA.Text')
                diem_b_vmix = get_field('DiemB.Text')
                sheet_row = rows[found_idx]
                ten_a_sheet = sheet_row.get(find_col_key(headers, 'VĐV A', 'TenA', 'Tên A'), '')
                ten_b_sheet = sheet_row.get(find_col_key(headers, 'VĐV B', 'TenB', 'Tên B'), '')
                a_sheet, b_sheet = last_word(ten_a_sheet), last_word(ten_b_sheet)
                a_vmix, b_vmix = last_word(ten_a_vmix), last_word(ten_b_vmix)
                if a_sheet and b_sheet and a_vmix and b_vmix:
                    if a_sheet == b_vmix and b_sheet == a_vmix:
                        diem_a, diem_b = diem_b_vmix, diem_a_vmix
                    else:
                        diem_a, diem_b = diem_a_vmix, diem_b_vmix
                else:
                    diem_a, diem_b = diem_a_vmix, diem_b_vmix
                if (a_sheet and b_sheet and a_vmix and b_vmix and not ((a_sheet == a_vmix and b_sheet == b_vmix) or (a_sheet == b_vmix and b_sheet == a_vmix))):
                    print(f"[CẢNH BÁO] Không xác định được mapping VĐV: SheetA={ten_a_sheet}, SheetB={ten_b_sheet}, vMixA={ten_a_vmix}, vMixB={ten_b_vmix}", file=sys.stderr)
                luotco = get_field('LuotCo.Text') or get_field('LuotCo')
                hr1 = get_field('HR1.Text') or get_field('HR1')
                hr2 = get_field('HR2.Text') or get_field('HR2')
                updates.append((found_idx, {
                    'Điểm A': diem_a,
                    'Điểm B': diem_b,
                    luotco_col: luotco,
                    hr1_col: hr1,
                    hr2_col: hr2
                }))
                # Save raw vMix values for later writing to AA..AL
                vmix_map[found_idx] = {
                    'TenA': ten_a_vmix,
                    'TenB': ten_b_vmix,
                    'DiemA': diem_a_vmix,
                    'DiemB': diem_b_vmix,
                    'Lco': luotco,
                    'HR1A': hr1,
                    'HR2A': hr2,
                    'HR1B': get_field('HR1B.Text') or get_field('HR1B'),
                    'HR2B': get_field('HR2B.Text') or get_field('HR2B'),
                    'AVGA': get_field('AvgA.Text') or get_field('AvgA'),
                    'AVGB': get_field('AvgB.Text') or get_field('AvgB'),
                }
            except Exception as ex:
                print(f'ERROR update_gsheet_with_vmix row {i}: {ex}', file=sys.stderr)
        # Chỉ update từng ô kết quả, không bao giờ ghi đè cả dòng/cột
        if not updates:
            self.status_var.set('Không có dòng nào được cập nhật!')
            return
        # Chỉ update các cột có tên đúng trong whitelist, tuyệt đối không update các cột công thức hoặc không liên quan
        allowed_names = [
            'Điểm A', 'Điểm B', 'Lượt cơ', 'HR1A', 'HR2A', 'HR1B', 'HR2B', diem_col, luotco_col, hr1_col, hr2_col
        ]
        updated_cells = 0
        for found_idx, update_dict in updates:
            row_num = found_idx + 2  # +2 vì header là dòng 1
            for col_name in allowed_names:
                if not col_name or col_name not in headers or col_name not in update_dict:
                    continue
                value = update_dict[col_name]
                old_value = rows[found_idx].get(col_name, '')
                def to_int(val):
                    try:
                        if val is None or val == '':
                            return ''
                        if isinstance(val, float):
                            return int(val)
                        if isinstance(val, int):
                            return val
                        if isinstance(val, str):
                            val = val.strip().replace("'", "")
                            if '.' in val:
                                return int(float(val))
                            return int(val)
                    except Exception:
                        return ''
                v = to_int(value)
                old_v = to_int(old_value)
                if v == old_v:
                    continue
                col_idx = headers.index(col_name)
                col_letter = chr(65 + col_idx)
                cell_range = f'Kết quả!{col_letter}{row_num}'
                print(f"[DEBUG] Update cell: {cell_range} (col: {col_name}) value: {v}")
                gs.write_table(cell_range, [[v]])
                updated_cells += 1
        # Sau khi ghi kết quả, tự động tính AVGA, AVGB nếu đủ dữ liệu
        for found_idx, update_row in updates:
            row_num = found_idx + 2  # +2 vì header là dòng 1
            # Lấy lại giá trị mới nhất (ưu tiên update_row, fallback sang rows)
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
                    if isinstance(val, str):
                        val = val.strip().replace("'", "").replace(",", ".")
                        return float(val)
                except Exception:
                    return None
            a = to_float(diem_a)
            b = to_float(diem_b)
            c = to_float(lco)
            avga = round(a/c, 3) if a is not None and c and c != 0 else ''
            avgb = round(b/c, 3) if b is not None and c and c != 0 else ''
            print(f"[DEBUG] AVGA: {avga}, AVGB: {avgb}, row: {row_num}, headers: {headers}")
            wrote = []
            for col_name, value in [('AVGA', avga), ('AVGB', avgb)]:
                if col_name in headers and value != '':
                    col_idx = headers.index(col_name)
                    col_letter = chr(65 + col_idx)
                    cell_range = f'Kết quả!{col_letter}{row_num}'
                    print(f"[DEBUG] Write {col_name} at {cell_range} value: {value:.3f}")
                    gs.write_table(cell_range, [[f"{value:.3f}"]])
                    wrote.append(f"{col_name}={value:.3f}")
            if wrote:
                import tkinter.messagebox as messagebox
                messagebox.showinfo('DEBUG', f'Đã ghi: {', '.join(wrote)} tại dòng {row_num}')
        self.status_var.set(f'Đã cập nhật {updated_cells} ô (Kết quả) lên Google Sheet!')

        # --- GHI BỔ SUNG: Ghi các trường vMix vào cột AA..AL (nếu ô rỗng) ---
        def index_to_col_letter(zero_based):
            n = zero_based + 1
            letters = ''
            while n > 0:
                n, rem = divmod(n-1, 26)
                letters = chr(65 + rem) + letters
            return letters
        extra_field_names = ['TenA','TenB','DiemA','DiemB','Lco','HR1A','HR2A','HR1B','HR2B','AVGA','AVGB']
        start_idx = 26  # AA
        for found_idx, vmix_vals in vmix_map.items():
            row_num = found_idx + 2
            for i, fname in enumerate(extra_field_names):
                col_idx = start_idx + i
                col_letter = index_to_col_letter(col_idx)
                cell_range = f'Kết quả!{col_letter}{row_num}'
                try:
                    existing = None
                    try:
                        existing = gs.read_table(cell_range)
                    except Exception:
                        existing = None
                    existing_val = ''
                    if existing:
                        if isinstance(existing, list) and existing and isinstance(existing[0], list):
                            existing_val = existing[0][0] if existing[0] else ''
                        elif isinstance(existing, list) and existing and isinstance(existing[0], dict):
                            vals = list(existing[0].values())
                            existing_val = vals[0] if vals else ''
                        elif isinstance(existing, str):
                            existing_val = existing
                        else:
                            existing_val = str(existing)
                    if str(existing_val).strip() == '':
                        val = vmix_vals.get(fname, '')
                        try:
                            import datetime
                            with open(os.path.join(os.getcwd(), 'vmix_write.log'), 'a', encoding='utf-8') as lf:
                                lf.write(f"[{datetime.datetime.now()}] WRITE {cell_range} => {val}\n")
                        except Exception:
                            pass
                        print(f"[DEBUG] WRITE {cell_range} => {val}")
                        gs.write_table(cell_range, [[val]])
                    else:
                        # preserve existing
                        try:
                            import datetime
                            with open(os.path.join(os.getcwd(), 'vmix_write.log'), 'a', encoding='utf-8') as lf:
                                lf.write(f"[{datetime.datetime.now()}] SKIP {cell_range} existing={existing_val}\n")
                        except Exception:
                            pass
                        print(f"[DEBUG] SKIP {cell_range} existing={existing_val}")
                except Exception as ex:
                    print(f'ERROR writing AA fields {fname} row {row_num}: {ex}', file=sys.stderr)

    def select_credentials(self):
        from tkinter import filedialog
        import os
        path = filedialog.askopenfilename(title='Chọn file credentials.json', filetypes=[('JSON files', '*.json')])
        if path:
            self.creds_path = path
            self.creds_label.config(text=os.path.basename(path), fg='#00FF00')
        else:
            self.creds_label.config(text='(Chưa chọn credentials)', fg='#FFD369')

    def __init__(self):
        super().__init__()
        self.base_title = 'QUẢN LÝ LIVESTREAM - NGUYỄN TIẾN CHUNG'
        self.title(self.base_title)
        self.after(100, lambda: self.animate_title(self, self.base_title))
        # Thêm icon/logo nếu có file logo.png
        import os, sys
        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.dirname(__file__)
            return os.path.join(base_path, relative_path)

        # Không dùng logo/icon cho chương trình
        self.configure(bg='#222831')
        self.state('zoomed')  # Full screen
        self.matches = []
        self.num_ban = 8
        self.sheet_url = ''
        self.sheet_rows = []  # raw rows from Google Sheets (for lookup by trận)
        self.creds_path = None
        self.ban_var = tk.IntVar(value=8)

        # Đường dẫn file lưu trạng thái tự động
        self._auto_state_path = os.path.join(os.path.dirname(__file__), '..', 'last_state.pkl')

        # Tự động khôi phục trạng thái nếu có
        self._auto_restore_state()

        self.create_widgets()

        # Sau khi tạo widget, khôi phục giá trị vào UI nếu có
        self._auto_restore_state_to_ui()

        # Gán sự kiện đóng cửa sổ để tự động lưu trạng thái
        self.protocol("WM_DELETE_WINDOW", self._on_close_and_save_state)

    def _auto_save_state(self):
        """Lưu trạng thái UI vào file khi đóng chương trình"""
        import pickle
        try:
            state = {
                'tengiai': self.tengiai_var.get() if hasattr(self, 'tengiai_var') else '',
                'thoigian': self.thoigian_var.get() if hasattr(self, 'thoigian_var') else '',
                'diadiem': self.diadiem_var.get() if hasattr(self, 'diadiem_var') else '',
                'chuchay': self.chuchay_var.get() if hasattr(self, 'chuchay_var') else '',
                'diemso': self.diemso_text.get('1.0', 'end-1c') if hasattr(self, 'diemso_text') else '',
                'sheet_url': self.url_var.get() if hasattr(self, 'url_var') else '',
                'creds_path': self.creds_path if hasattr(self, 'creds_path') else '',
                'ban': self.ban_var.get() if hasattr(self, 'ban_var') else 8,
                'table': [
                    [row[i].get() for i in range(6)]
                    for row in getattr(self, 'match_rows', [])
                ] if hasattr(self, 'match_rows') else []
            }
            with open(self._auto_state_path, 'wb') as f:
                pickle.dump(state, f)
        except Exception as ex:
            print(f"[WARN] Không thể lưu trạng thái tự động: {ex}")

    def _auto_restore_state(self):
        """Đọc trạng thái từ file, lưu vào self._restored_state"""
        import pickle
        self._restored_state = None
        try:
            if os.path.exists(self._auto_state_path):
                with open(self._auto_state_path, 'rb') as f:
                    self._restored_state = pickle.load(f)
        except Exception as ex:
            print(f"[WARN] Không thể khôi phục trạng thái tự động: {ex}")

    def _auto_restore_state_to_ui(self):
        """Gán lại giá trị vào UI từ self._restored_state nếu có"""
        s = getattr(self, '_restored_state', None)
        if not s:
            return
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

    def create_widgets(self):



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
            if hasattr(self, 'match_rows'):
                for row in self.match_rows:
                    diem_entry = row[4]
                    if isinstance(diem_entry, tk.Entry):
                        diem_entry.delete(0, 'end')
                        diem_entry.insert(0, value)
        self.diemso_text.bind('<Return>', lambda e: (update_all_scores(), 'break'))
        self.diemso_text.bind('<Control-Return>', lambda e: (update_all_scores(), 'break'))
        # Row 1: ĐỊA ĐIỂM
        tk.Label(header_frame, text='Địa điểm:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=1, column=col_offset+0, sticky='e', padx=(10,4), pady=6)
        self.diadiem_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.diadiem_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369').grid(row=1, column=col_offset+1, columnspan=5, sticky='ew', padx=(4,10), pady=6)
        # Row 2: CHỮ CHẠY
        tk.Label(header_frame, text='Chữ chạy:', fg='#FFD369', bg='#1A2233', font=('Segoe UI', 18, 'bold')).grid(row=2, column=col_offset+0, sticky='e', padx=(10,4), pady=6)
        self.chuchay_var = tk.StringVar()
        tk.Entry(header_frame, textvariable=self.chuchay_var, font=('Segoe UI', 18), relief='groove', bd=3, bg='#232B3E', fg='white', insertbackground='white', highlightthickness=1, highlightbackground='#FFD369').grid(row=2, column=col_offset+1, columnspan=5, sticky='ew', padx=(4,10), pady=6)

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
            self.animate_title(win, win.base_title)
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
        # Di chuyển nút Lưu bảng và Tải bảng xuống thanh trạng thái dưới đáy màn hình
        # (Tạo và pack chỉ ở đây, sau khi đã có bottom_frame)

        # Nút preview tổng hợp bảng điểm
        def open_preview_all():
            import requests
            import xml.etree.ElementTree as ET
            preview = tk.Toplevel(self)
            preview.base_title = 'Preview tổng hợp bảng điểm các bàn'
            preview.title(preview.base_title)
            self.animate_title(preview, preview.base_title)
            num_ban = self.ban_var.get()
            cols = 4 if num_ban > 4 else num_ban
            rows = (num_ban + cols - 1) // cols
            preview.frames = []

            auto_update = {'enabled': False, 'job': None}

            def fetch_all_vmix():
                for i in range(num_ban):
                    try:
                        row = self.match_rows[i]
                        vmix_url = row[5].get().strip() if hasattr(row[5], 'get') else ''
                        if not vmix_url:
                            continue
                        # Lấy XML từ vMix
                        resp = requests.get(f'{vmix_url}/API/', timeout=2)
                        xml = resp.text
                        root = ET.fromstring(xml)
                        # Lấy Input 1
                        input1 = root.find(".//input[@number='1']")
                        # Lấy các field
                        def get_field(name):
                            if input1 is not None:
                                for text in input1.findall('text'):
                                    if text.attrib.get('name') == name:
                                        return text.text or ''
                            return ''
                        ten_a = get_field('TenA.Text')
                        ten_b = get_field('TenB.Text')
                        tran = get_field('Tran.Text')
                        diem = get_field('Noi dung.Text')
                        diem_a = get_field('DiemA.Text') or get_field('DiemA')
                        diem_b = get_field('DiemB.Text') or get_field('DiemB')
                        # Lấy thông số riêng cho A và B nếu có
                        # Lấy đúng AvgA.Text, AvgB.Text, Lco.Text
                        avg_a = get_field('AvgA.Text') or get_field('AVGA.Text') or get_field('AVGA') or ''
                        avg_b = get_field('AvgB.Text') or get_field('AVGB.Text') or get_field('AVGB') or ''
                        lco = get_field('Lco.Text') or get_field('Lco') or ''
                        # Các trường khác giữ nguyên
                        luot_co = get_field('LuotCo.Text') or get_field('LuotCo') or ''
                        luot_co_a = get_field('LuotCoA.Text') or get_field('LuotCoA') or ''
                        luot_co_b = get_field('LuotCoB.Text') or get_field('LuotCoB') or ''
                        hr1_a = get_field('HR1A.Text') or get_field('HR1A') or ''
                        hr2_a = get_field('HR2A.Text') or get_field('HR2A') or ''
                        hr1_b = get_field('HR1B.Text') or get_field('HR1B') or ''
                        hr2_b = get_field('HR2B.Text') or get_field('HR2B') or ''
                        # Nếu không có riêng thì lấy chung
                        hr1 = get_field('HR1.Text') or get_field('HR1') or ''
                        hr2 = get_field('HR2.Text') or get_field('HR2') or ''
                        # Cập nhật label trong frame
                        frame = preview.frames[i]
                        for w in frame.winfo_children():
                            w.destroy()
                        ban = row[1].get() if hasattr(row[1], 'get') else f'BÀN {i+1:02d}'
                        status = row[-1].cget('text') if hasattr(row[-1], 'cget') else ''
                        # Layout: tên, điểm, các thông số chi tiết từng VĐV
                        # Dòng 1: Tên VĐV A (trắng) - Tên VĐV B (vàng)
                        row1 = tk.Frame(frame, bg='#111')
                        row1.pack(fill='x', padx=4, pady=(2,0))
                        tk.Label(row1, text=ten_a, font=('Arial', 18, 'bold'), fg='white', bg='#111', anchor='w').pack(side='left', padx=(0,10), fill='x', expand=True)
                        tk.Label(row1, text=ban, font=('Arial', 14, 'bold'), fg='#FFD369', bg='#111', anchor='center').pack(side='left', padx=(0,10))
                        tk.Label(row1, text=ten_b, font=('Arial', 18, 'bold'), fg='#FFD600', bg='#111', anchor='e').pack(side='left', fill='x', expand=True)

                        # Dòng 2: Điểm số lớn (trắng - vàng)
                        row2 = tk.Frame(frame, bg='#111')
                        row2.pack(fill='x', padx=4, pady=(0,2))
                        diem_a_str = diem_a if diem_a is not None else '-'
                        diem_b_str = diem_b if diem_b is not None else '-'
                        tk.Label(row2, text=diem_a_str, font=('Arial', 48, 'bold'), fg='white', bg='#111').pack(side='left', padx=(0,10), fill='x', expand=True)
                        tk.Label(row2, text=f'TRẬN {tran}', font=('Arial', 22, 'bold'), fg='#00E5FF', bg='#111').pack(side='left', padx=(0,10))
                        tk.Label(row2, text=diem_b_str, font=('Arial', 48, 'bold'), fg='#FFD600', bg='#111').pack(side='left', fill='x', expand=True)

                        # Dòng 3: AVG, Lco, HR1, HR2 từng bên
                        row3 = tk.Frame(frame, bg='#111')
                        row3.pack(fill='x', padx=4, pady=(0,2))
                        tk.Label(row3, text=f"AVG", font=('Arial', 16, 'bold'), fg='white', bg='#111').pack(side='left', padx=(0,2))
                        tk.Label(row3, text=f"{avg_a or '-'}", font=('Arial', 16, 'bold'), fg='white', bg='#111').pack(side='left', padx=(0,8))
                        tk.Label(row3, text=f"Lco", font=('Arial', 16, 'bold'), fg='#FFD369', bg='#111').pack(side='left', padx=(0,2))
                        tk.Label(row3, text=f"{lco or '-'}", font=('Arial', 16, 'bold'), fg='#FFD369', bg='#111').pack(side='left', padx=(0,8))
                        tk.Label(row3, text=f"AVG", font=('Arial', 16, 'bold'), fg='#FFD600', bg='#111').pack(side='left', padx=(0,2))
                        tk.Label(row3, text=f"{avg_b or '-'}", font=('Arial', 16, 'bold'), fg='#FFD600', bg='#111').pack(side='left', padx=(0,8))

                        # Dòng 4: HR1, HR2 từng bên
                        row4 = tk.Frame(frame, bg='#111')
                        row4.pack(fill='x', padx=4, pady=(0,2))
                        tk.Label(row4, text=f"HR1", font=('Arial', 14, 'bold'), fg='white', bg='#111').pack(side='left', padx=(0,2))
                        tk.Label(row4, text=f"{hr1_a or hr1 or '-'}", font=('Arial', 14, 'bold'), fg='white', bg='#111').pack(side='left', padx=(0,8))
                        tk.Label(row4, text=f"HR2", font=('Arial', 14, 'bold'), fg='white', bg='#111').pack(side='left', padx=(0,2))
                        tk.Label(row4, text=f"{hr2_a or hr2 or '-'}", font=('Arial', 14, 'bold'), fg='white', bg='#111').pack(side='left', padx=(0,8))
                        tk.Label(row4, text=f"HR1", font=('Arial', 14, 'bold'), fg='#FFD600', bg='#111').pack(side='left', padx=(0,2))
                        tk.Label(row4, text=f"{hr1_b or hr1 or '-'}", font=('Arial', 14, 'bold'), fg='#FFD600', bg='#111').pack(side='left', padx=(0,8))
                        tk.Label(row4, text=f"HR2", font=('Arial', 14, 'bold'), fg='#FFD600', bg='#111').pack(side='left', padx=(0,2))
                        tk.Label(row4, text=f"{hr2_b or hr2 or '-'}", font=('Arial', 14, 'bold'), fg='#FFD600', bg='#111').pack(side='left', padx=(0,8))

                        # Dòng trạng thái
                        tk.Label(frame, text=f'TRẠNG THÁI: {status}', font=('Arial', 11, 'italic'), fg='#FFD369', bg='#111').pack(anchor='w', padx=4, pady=(0,2))
                        tk.Label(frame, text=f'TRẠNG THÁI: {status}', font=('Arial', 11, 'italic'), fg='#FFD369', bg='#111').pack(anchor='w', padx=4, pady=(0,2))
                    except Exception as ex:
                        frame = preview.frames[i]
                        for w in frame.winfo_children():
                            w.destroy()
                        tk.Label(frame, text=f'Lỗi lấy vMix: {ex}', font=('Arial', 12, 'italic'), fg='red', bg='#111').pack(anchor='w', padx=4, pady=4)

            for i in range(num_ban):
                r, c = divmod(i, cols)
                frame = tk.Frame(preview, bg='#111', bd=2, relief='groove')
                frame.grid(row=r, column=c, padx=6, pady=6, sticky='nsew')
                preview.frames.append(frame)
                # Hiển thị dữ liệu ban đầu (từ bảng)
                try:
                    row = self.match_rows[i]
                    ten_a = row[2].get() if hasattr(row[2], 'get') else ''
                    ten_b = row[3].get() if hasattr(row[3], 'get') else ''
                    diem_a = row[4].get() if hasattr(row[4], 'get') else ''
                    tran = row[0].get() if hasattr(row[0], 'get') else ''
                    ban = row[1].get() if hasattr(row[1], 'get') else f'BÀN {i+1:02d}'
                    status = row[-1].cget('text') if hasattr(row[-1], 'cget') else ''
                except Exception:
                    ten_a = ten_b = diem_a = tran = ban = status = ''
                tk.Label(frame, text=ban, font=('Arial', 16, 'bold'), fg='#FFD369', bg='#111').pack(anchor='w', padx=4, pady=(2,0))
                tk.Label(frame, text=f'TRẬN: {tran}', font=('Arial', 14, 'bold'), fg='#00E5FF', bg='#111').pack(anchor='w', padx=4)
                tk.Label(frame, text=f'{ten_a}  vs  {ten_b}', font=('Arial', 13, 'bold'), fg='#fff', bg='#111').pack(anchor='w', padx=4)
                tk.Label(frame, text=f'ĐIỂM: {diem_a}', font=('Arial', 20, 'bold'), fg='#00FF00', bg='#111').pack(anchor='w', padx=4, pady=(0,2))
                tk.Label(frame, text=f'TRẠNG THÁI: {status}', font=('Arial', 11, 'italic'), fg='#FFD369', bg='#111').pack(anchor='w', padx=4, pady=(0,2))


            # Hàm tự động cập nhật
            def auto_update_vmix():
                if auto_update['enabled']:
                    fetch_all_vmix()
                    # Lưu job để có thể hủy
                    auto_update['job'] = preview.after(5000, auto_update_vmix)

            def on_toggle_auto():
                auto_update['enabled'] = bool(auto_var.get())
                if auto_update['enabled']:
                    fetch_all_vmix()
                    auto_update['job'] = preview.after(5000, auto_update_vmix)
                else:
                    if auto_update['job']:
                        preview.after_cancel(auto_update['job'])
                        auto_update['job'] = None

            # Nút lấy dữ liệu và checkbox tự động
            control_frame = tk.Frame(preview, bg='#232B3E')
            control_frame.grid(row=rows, column=0, columnspan=cols, sticky='ew', padx=10, pady=10)
            btn = tk.Button(control_frame, text='Lấy dữ liệu từ vMix', font=('Arial', 13, 'bold'), bg='#FFD369', fg='#222831', command=fetch_all_vmix)
            btn.pack(side='left', padx=4)
            auto_var = tk.IntVar(value=0)
            auto_check = tk.Checkbutton(control_frame, text='Tự động cập nhật (5s)', variable=auto_var, command=on_toggle_auto, font=('Arial', 12), bg='#232B3E', fg='#FFD369', selectcolor='#232B3E', activebackground='#232B3E', activeforeground='#FFD369')
            auto_check.pack(side='left', padx=8)
            # Khi đóng cửa sổ preview thì hủy auto-update nếu đang chạy
            def on_close():
                auto_update['enabled'] = False
                if auto_update['job']:
                    preview.after_cancel(auto_update['job'])
                    auto_update['job'] = None
                preview.destroy()
            preview.protocol('WM_DELETE_WINDOW', on_close)

            for c in range(cols):
                preview.grid_columnconfigure(c, weight=1)
            for r in range(rows+1):
                preview.grid_rowconfigure(r, weight=1)

        # ...existing code...
        # Sau khi tạo bottom_frame, pack status_label phía dưới
        # Đặt đoạn này vào cuối hàm create_widgets
            # Địa chỉ vMix input
            # Địa chỉ vMix đã chuyển xuống từng dòng, không cần ô nhập chung
            # Số bàn input và bộ lọc tìm kiếm chuyển xuống dưới cùng
            bottom_frame = tk.Frame(self, bg='#222831')
            bottom_frame.pack(fill='x', pady=5, side='bottom')
            tk.Label(bottom_frame, text='Tổng số bàn:', fg='#FFD369', bg='#222831', font=('Arial', 12, 'bold')).pack(side='left', padx=10)
            ban_spin = tk.Spinbox(bottom_frame, from_=1, to=32, textvariable=self.ban_var, width=5, font=('Arial', 12))
            ban_spin.pack(side='left', padx=5)
            tk.Button(bottom_frame, text='Cập nhật', command=self.reload_matches, bg='#FFD369', fg='#222831', font=('Arial', 11, 'bold')).pack(side='left', padx=10)
            tk.Label(bottom_frame, text='Lọc bàn:', fg='#FFD369', bg='#222831', font=('Arial', 11)).pack(side='left', padx=5)
            self.filter_ban_var = tk.StringVar()
            self.filter_ban_entry = tk.Entry(bottom_frame, textvariable=self.filter_ban_var, width=8, font=('Arial', 11))
            self.filter_ban_entry.pack(side='left', padx=2)
            tk.Label(bottom_frame, text='Tìm VĐV:', fg='#FFD369', bg='#222831', font=('Arial', 11)).pack(side='left', padx=5)
            self.filter_vdv_var = tk.StringVar()
            self.filter_vdv_entry = tk.Entry(bottom_frame, textvariable=self.filter_vdv_var, width=16, font=('Arial', 11))
            self.filter_vdv_entry.pack(side='left', padx=2)
            tk.Button(bottom_frame, text='Lọc', command=self.populate_table, bg='#FFD369', fg='#222831', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
            # Đưa nút 'Preview tổng hợp' xuống bottom_frame, gần góc phải
            self.preview_all_btn = tk.Button(bottom_frame, text='Preview tổng hợp', command=open_preview_all, bg='#00B8D4', fg='#fff', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#4DD0E1', activeforeground='#222831')
            self.preview_all_btn.pack(side='right', padx=10)
            # Thêm dòng báo trạng thái bên phải, nổi bật màu đỏ khi lỗi
            self.status_label = tk.Label(bottom_frame, textvariable=self.status_var, fg='#FF5252', bg='#222831', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2)
            self.status_label.pack(side='right', padx=10)
            self.table_frame.bind('<Configure>', self._on_table_configure)
            self.table_canvas.bind('<Configure>', self._on_canvas_configure)
        self.status_var = tk.StringVar(value='')

        # Table with scrollbars (Entry-based)
        table_outer = tk.Frame(self, bg='#1A2233', highlightbackground='#232B3E', highlightthickness=2)
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

        # Chia tỉ lệ % chiều ngang cho các cột
        # 8 cột: Trận, Số bàn, Tên VĐV A, Tên VĐV B, Điểm số, Địa chỉ vMix, Kết quả, Gửi
        # Chia lại: Trận 2%, BÀN 10%, các cột còn lại chia lại cho đủ 100%
        # 8 cột: Trận, BÀN, Tên VĐV A, Tên VĐV B, Điểm số, Địa chỉ vMix, Kết quả, Gửi
        # Phân bổ lại: Trận 14%, BÀN 8%, Tên A 15%, Tên B 15%, Điểm số 6%, Địa chỉ vMix 21%, Kết quả 6%, Gửi 6%, Sửa 6%
        col_weights = [12, 8, 15, 15, 4, 21, 2, 2, 2]
        total_weight = sum(col_weights)
        for col, weight in enumerate(col_weights):
            self.table_frame.grid_columnconfigure(col, weight=weight)
        self.match_rows = []
        self.populate_table()

        # Địa chỉ vMix input
        # Địa chỉ vMix đã chuyển xuống từng dòng, không cần ô nhập chung
        # Số bàn input và bộ lọc tìm kiếm chuyển xuống dưới cùng
        bottom_frame = tk.Frame(self, bg='#222831')
        bottom_frame.pack(fill='x', pady=5, side='bottom')
        tk.Label(bottom_frame, text='Tổng số bàn:', fg='#FFD369', bg='#222831', font=('Arial', 12, 'bold')).pack(side='left', padx=10)
        ban_spin = tk.Spinbox(bottom_frame, from_=1, to=32, textvariable=self.ban_var, width=5, font=('Arial', 12))
        ban_spin.pack(side='left', padx=5)
        # (Đoạn này đã được xử lý ở phía dưới, không cần pack ở đây)
        tk.Button(bottom_frame, text='Cập nhật', command=self.reload_matches, bg='#FFD369', fg='#222831', font=('Arial', 11, 'bold')).pack(side='left', padx=10)
        tk.Label(bottom_frame, text='Lọc bàn:', fg='#FFD369', bg='#222831', font=('Arial', 11)).pack(side='left', padx=5)
        self.filter_ban_var = tk.StringVar()
        self.filter_ban_entry = tk.Entry(bottom_frame, textvariable=self.filter_ban_var, width=8, font=('Arial', 11))
        self.filter_ban_entry.pack(side='left', padx=2)
        tk.Label(bottom_frame, text='Tìm VĐV:', fg='#FFD369', bg='#222831', font=('Arial', 11)).pack(side='left', padx=5)
        self.filter_vdv_var = tk.StringVar()
        self.filter_vdv_entry = tk.Entry(bottom_frame, textvariable=self.filter_vdv_var, width=16, font=('Arial', 11))
        self.filter_vdv_entry.pack(side='left', padx=2)
        tk.Button(bottom_frame, text='Lọc', command=self.populate_table, bg='#FFD369', fg='#222831', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        # Nút 'Preview tổng hợp' đặt gần góc phải, trước dòng trạng thái
        self.preview_all_btn = tk.Button(bottom_frame, text='Preview tổng hợp', command=open_preview_all, bg='#00B8D4', fg='#fff', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#4DD0E1', activeforeground='#222831')
        self.preview_all_btn.pack(side='right', padx=10)
        # Thêm dòng báo trạng thái bên phải
        self.status_label = tk.Label(bottom_frame, textvariable=self.status_var, fg='#00FF00', bg='#222831', font=('Segoe UI', 12, 'italic'))
        self.status_label.pack(side='right', padx=10)
        # Nút lấy tất cả từ vMix
        btn_fetch_all = tk.Button(bottom_frame, text='Lấy tất cả từ vMix', command=self.fetch_all_vmix_to_table, bg='#00C853', fg='#222831', font=('Arial', 11, 'bold'))
        btn_fetch_all.pack(side='right', padx=10)
        # Thêm nút Lưu bảng và Tải bảng xuống dưới cùng (sau khi đã có bottom_frame)
        self.save_btn = tk.Button(bottom_frame, text='Lưu bảng', command=self.save_table_to_csv, bg='#00C853', fg='#1A2233', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#B9F6CA', activeforeground='#1A2233')
        self.load_btn = tk.Button(bottom_frame, text='Tải bảng', command=self.load_table_from_csv, bg='#2962FF', fg='#FFD369', font=('Segoe UI', 12, 'bold'), relief='groove', bd=2, activebackground='#82B1FF', activeforeground='#FFD369')
        self.save_btn.pack(in_=bottom_frame, side='right', padx=10)
        self.load_btn.pack(in_=bottom_frame, side='right', padx=5)
        self.table_frame.bind('<Configure>', self._on_table_configure)
        self.table_canvas.bind('<Configure>', self._on_canvas_configure)

    def _on_table_configure(self, event):
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        # Make inner frame width match canvas width if possible
        canvas_width = event.width
        self.table_canvas.itemconfig(self.table_window, width=canvas_width)

    def clear_table(self):
        for widgets in self.match_rows:
            for w in widgets:
                w.destroy()
        self.match_rows.clear()

    def populate_table(self):
        # Lưu lại toàn bộ giá trị các ô cũ nếu có
        old_rows = []
        if hasattr(self, 'match_rows') and self.match_rows:
            for row in self.match_rows:
                # [Trận, Bàn, Tên A, Tên B, Điểm, vMix]
                vals = []
                for idx in range(6):
                    if idx < len(row):
                        vals.append(row[idx].get() if hasattr(row[idx], 'get') else '')
                    else:
                        vals.append('')
                old_rows.append(vals)
        self.clear_table()
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

            # Số trận (tăng width thêm cho dễ nhìn, sticky nsew)
            # Tăng width, ipadx, font cho cột Trận (bên trái)
            # Giảm width, font, ipadx cho Entry cột Trận (60% so với trước)
            e_tran = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), relief='groove', bd=2, justify='center', insertbackground='white', width=4)
            if i < len(old_rows):
                e_tran.insert(0, old_rows[i][0])
            e_tran.grid(row=i+1, column=0, padx=2, pady=2, ipadx=2, ipady=6, sticky='nsew')
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
            e_ban = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), relief='groove', bd=2, insertbackground='white', justify='center', width=5)
            if i < len(old_rows):
                e_ban.insert(0, old_rows[i][1])
            e_ban.grid(row=i+1, column=1, padx=2, pady=2, ipadx=2, ipady=6, sticky='ew')
            # Hiệu ứng dấu nháy trắng khi focus
            def on_focus_in_ban(event, entry=e_ban):
                try:
                    entry.config(insertbackground='white')
                except:
                    pass
            e_ban.bind('<FocusIn>', on_focus_in_ban)
            widgets.append(e_ban)
            # Tên VĐV A
            e_a = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg='#222831', relief='groove', bd=2, width=16)
            if i < len(old_rows):
                e_a.insert(0, old_rows[i][2])
            e_a.config(state='readonly', fg='#222831')
            e_a.grid(row=i+1, column=2, padx=2, pady=2, ipadx=6, ipady=6, sticky='ew')
            widgets.append(e_a)
            # Tên VĐV B
            e_b = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg='#222831', relief='groove', bd=2, width=16)
            if i < len(old_rows):
                e_b.insert(0, old_rows[i][3])
            e_b.config(state='readonly', fg='#222831')
            e_b.grid(row=i+1, column=3, padx=2, pady=2, ipadx=6, ipady=6, sticky='ew')
            widgets.append(e_b)
            # Điểm số (giảm width)
            e_diem = tk.Entry(self.table_frame, font=('Arial', 18), bg=bg, fg=contrast_fg(bg), justify='center', relief='groove', bd=2, width=5)
            if i < len(old_rows):
                e_diem.insert(0, old_rows[i][4])
            e_diem.config(width=13)
            e_diem.grid(row=i+1, column=4, padx=2, pady=2, ipadx=7, ipady=6, sticky='ew')
            def diem_user_edit(event, entry=e_diem):
                entry._user_edited = True
            e_diem.bind('<Key>', diem_user_edit)
            widgets.append(e_diem)
            # Địa chỉ vMix cho từng dòng (giữ lại giá trị cũ nếu có)
            e_vmix = tk.Entry(self.table_frame, font=('Arial', 16), bg='#B3E5FC', fg='#222831', relief='groove', bd=2, width=22)
            if i < len(old_rows) and old_rows[i][5]:
                e_vmix.insert(0, old_rows[i][5])
            else:
                e_vmix.insert(0, 'http://127.0.0.1:8088')
            e_vmix.config(width=9)
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
                    write_range = f'Kết quả!A1:{chr(65+len(headers)-1)}{len(new_values)}'
                    gs.write_table(write_range, new_values)

                    # Sau khi ghi kết quả, tự động tính AVGA, AVGB nếu đủ dữ liệu
                    # Lấy lại giá trị mới nhất (ưu tiên update_row, fallback sang rows)
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
                    wrote = []
                    for col_name, value in [('AVGA', avga), ('AVGB', avgb)]:
                        if col_name in headers and value != '':
                            col_idx = headers.index(col_name)
                            col_letter = chr(65 + col_idx)
                            cell_range = f'Kết quả!{col_letter}{found_idx+2}'
                            # Định dạng số kiểu 0,000 (dấu phẩy)
                            value_str = f"{value:.3f}".replace('.', ',')
                            gs.write_table(cell_range, [[value_str]])
                            wrote.append(f"{col_name}={value_str}")
                    messagebox.showinfo('Thành công', f'Đã cập nhật kết quả trận {tran_val} lên Google Sheet!{'\n' if wrote else ''}{'\n'.join(wrote)}')
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
        keys = list(self.sheet_rows[0].keys())
        print(f'DEBUG: sheet_rows[0] keys = {keys}', file=sys.stderr)
        # Cho phép người dùng chỉnh tên cột tại đây:
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
        for idx, r in enumerate(self.sheet_rows):
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
        self.populate_table()
        # Nếu không có dữ liệu, xóa tên VĐV và trạng thái từng dòng
        if not self.sheet_rows:
            for row in self.match_rows:
                if len(row) >= 6:
                    row[2].config(state='normal', fg='#222831'); row[2].delete(0, 'end'); row[2].config(state='readonly', fg='#222831')
                    row[3].config(state='normal', fg='#222831'); row[3].delete(0, 'end'); row[3].config(state='readonly', fg='#222831')
                    row[-1].config(text='', fg='#005A9E')

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
