import os, sys
sys.path.insert(0, os.path.abspath('src'))

# Create a fake tkinter module to avoid opening a real GUI during test
import types
fake_tk = types.SimpleNamespace()

class FakeToplevel:
    def __init__(self, parent=None):
        print('FakeToplevel.__init__')
        self._props = {}
    def title(self, t): print('FakeToplevel.title', t)
    def geometry(self, g): print('FakeToplevel.geometry', g)
    def grab_set(self): pass
    def transient(self, p): pass
    def destroy(self): print('FakeToplevel.destroy')
    def configure(self, **kwargs):
        self._props.update(kwargs)
    def __getitem__(self, key):
        return self._props.get(key, None)

class FakeLabel:
    def __init__(self, *args, **kwargs): pass
    def place(self, *a, **k): pass

class FakeEntryW:
    def __init__(self, *args, **kwargs):
        self._val = ''
    def insert(self, idx, val):
        self._val = val
    def place(self, *a, **k): pass
    def get(self):
        return self._val

captured_commands = []
class FakeButton:
    def __init__(self, *a, **k):
        cmd = k.get('command') or (a[1] if len(a) > 1 and callable(a[1]) else None)
        if cmd:
            captured_commands.append(cmd)
    def place(self, *a, **k): pass

class FakeMB:
    @staticmethod
    def showerror(title, msg): print('MB.showerror:', title, msg)
    @staticmethod
    def showinfo(title, msg): print('MB.showinfo:', title, msg)

fake_tk.Toplevel = FakeToplevel
fake_tk.Label = FakeLabel
fake_tk.Entry = FakeEntryW
fake_tk.Button = FakeButton
fake_tk.messagebox = FakeMB
import types as _types
fake_tk.ttk = _types.SimpleNamespace()
fake_tk.simpledialog = _types.SimpleNamespace()
class FakeRoot:
    def __init__(self): pass
    def title(self, *a, **k): pass
    def withdraw(self): pass
    def mainloop(self): pass

fake_tk.Tk = FakeRoot

# Install into sys.modules so `import tkinter as tk` inside the target function uses this fake
sys.modules['tkinter'] = fake_tk
sys.modules['tkinter.messagebox'] = FakeMB

# Import target module and patch GSheetClient with dummy implementation
import importlib
import gui_fullscreen_match as gm

class DummyGSheetClient:
    def __init__(self, sid, creds):
        print('DummyGSheetClient.init', sid, creds)
    def read_table(self, rng):
        print('DummyGSheetClient.read_table', rng)
        return [{
            'Trận': '1', 'Tên VĐV A': 'Alice', 'Tên VĐV B': 'Bob',
            'Điểm A': '10','Điểm B': '9','Lượt cơ': '3',
            'HR1A':'','HR2A':'','HR1B':'','HR2B':'','AVGA':'7','AVGB':'8'
        }]
    def batch_update(self, batch):
        print('DummyGSheetClient.batch_update called with', batch)

gm.GSheetClient = DummyGSheetClient

# Ensure dummy creds file exists
os.makedirs('tools', exist_ok=True)
creds = os.path.join('tools', 'dummy_creds.json')
with open(creds, 'w', encoding='utf-8') as f:
    f.write('{}')

# Prepare minimal fake instance of FullScreenMatchGUI
class SimpleGet:
    def __init__(self, v): self._v = v
    def get(self): return self._v

obj = gm.FullScreenMatchGUI.__new__(gm.FullScreenMatchGUI)
obj.url_var = type('X', (), {'get': lambda self: 'https://docs.google.com/spreadsheets/d/FAKEID/edit'})()
obj.creds_path = creds
obj.match_rows = [(SimpleGet('1'),)]

# Run the popup function (uses fake tkinter and dummy sheet client)
print('Calling open_edit_popup...')
obj.open_edit_popup(0)
print('open_edit_popup finished')
# simulate pressing the first captured button (should be 'Lưu')
if captured_commands:
    print('Invoking captured command (simulating Save)...')
    try:
        captured_commands[0]()
    except Exception as e:
        print('Error invoking captured command:', e)
    print('After invoking captured command')
else:
    print('No button commands captured')
