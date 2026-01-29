import tkinter as tk
from tkinter import ttk, messagebox

# Dummy data for demo (replace with Google Sheets fetch)
MATCHES = [
    {
        'ban': 'BÀN 01',
        'tran': 'TRẬN 1',
        'vdv_a': 'QUANG HIỆP (Vũ Ngọc)',
        'vdv_b': 'TÂN HUYỀN (Tony Fruit)',
        'info': 'CÚP 3C BLACKPINK LẦN 2 - 2026',
    },
    {
        'ban': 'BÀN 02',
        'tran': 'TRẬN 2',
        'vdv_a': 'PHẠM HẬU (Blackpink)',
        'vdv_b': 'HỮU NHÂN (Subin)',
        'info': 'CÚP 3C BLACKPINK LẦN 2 - 2026',
    },
    # ... add more matches as needed
]

class MatchSelectorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Quản lý Trận đấu - vMix Scoreboard')
        self.geometry('700x300')
        self.matches = MATCHES
        self.create_widgets()

    def create_widgets(self):
        # Dropdown chọn bàn
        tk.Label(self, text='Chọn Bàn:', width=16, anchor='e').grid(row=0, column=0, sticky='e', padx=(12,4))
        self.ban_var = tk.StringVar()
        ban_list = sorted(set(m['ban'] for m in self.matches))
        self.ban_cb = ttk.Combobox(self, textvariable=self.ban_var, values=ban_list, state='readonly', width=24)
        self.ban_cb.grid(row=0, column=1, sticky='w', padx=(0,12))
        self.ban_cb.bind('<<ComboboxSelected>>', self.update_tran_options)

        # Dropdown chọn trận
        tk.Label(self, text='Chọn Trận:', width=16, anchor='e').grid(row=1, column=0, sticky='e', padx=(12,4))
        self.tran_var = tk.StringVar()
        self.tran_cb = ttk.Combobox(self, textvariable=self.tran_var, values=[], state='readonly', width=24)
        self.tran_cb.grid(row=1, column=1, sticky='w', padx=(0,12))
        self.tran_cb.bind('<<ComboboxSelected>>', self.display_match_info)

        # Thông tin VĐV và trận
        self.info_frame = tk.Frame(self)
        self.info_frame.grid(row=2, column=0, columnspan=2, sticky='w', pady=10)
        self.vdv_a_label = tk.Label(self.info_frame, text='VĐV A:', width=16, anchor='e')
        self.vdv_a_label.grid(row=0, column=0, sticky='e', padx=(12,4))
        self.vdv_b_label = tk.Label(self.info_frame, text='VĐV B:', width=16, anchor='e')
        self.vdv_b_label.grid(row=1, column=0, sticky='e', padx=(12,4))
        self.info_label = tk.Label(self.info_frame, text='Thông tin trận:', width=16, anchor='e')
        self.info_label.grid(row=2, column=0, sticky='e', padx=(12,4))

        self.vdv_a_val = tk.Label(self.info_frame, text='', width=32, anchor='w')
        self.vdv_a_val.grid(row=0, column=1, sticky='w', padx=(0,12))
        self.vdv_b_val = tk.Label(self.info_frame, text='', width=32, anchor='w')
        self.vdv_b_val.grid(row=1, column=1, sticky='w', padx=(0,12))
        self.info_val = tk.Label(self.info_frame, text='', width=32, anchor='w')
        self.info_val.grid(row=2, column=1, sticky='w', padx=(0,12))

        # Nút gửi lên vMix
        self.push_btn = tk.Button(self, text='Gửi lên vMix', command=self.push_to_vmix)
        self.push_btn.grid(row=3, column=0, columnspan=2, pady=10)

    def update_tran_options(self, event=None):
        ban = self.ban_var.get()
        tran_list = [m['tran'] for m in self.matches if m['ban'] == ban]
        self.tran_cb['values'] = tran_list
        if tran_list:
            self.tran_var.set(tran_list[0])
            self.display_match_info()
        else:
            self.tran_var.set('')
            self.display_match_info()

    def display_match_info(self, event=None):
        ban = self.ban_var.get()
        tran = self.tran_var.get()
        match = next((m for m in self.matches if m['ban'] == ban and m['tran'] == tran), None)
        if match:
            self.vdv_a_val.config(text=match['vdv_a'])
            self.vdv_b_val.config(text=match['vdv_b'])
            self.info_val.config(text=match['info'])
        else:
            self.vdv_a_val.config(text='')
            self.vdv_b_val.config(text='')
            self.info_val.config(text='')

    def push_to_vmix(self):
        ban = self.ban_var.get()
        tran = self.tran_var.get()
        match = next((m for m in self.matches if m['ban'] == ban and m['tran'] == tran), None)
        if not match:
            messagebox.showwarning('Chưa chọn trận', 'Vui lòng chọn bàn và trận hợp lệ!')
            return
        # TODO: Gọi hàm gửi dữ liệu lên vMix ở đây
        messagebox.showinfo('Đã gửi', f"Đã gửi trận {tran} ({ban}) lên vMix!\nVĐV A: {match['vdv_a']}\nVĐV B: {match['vdv_b']}")

if __name__ == '__main__':
    app = MatchSelectorApp()
    app.mainloop()
