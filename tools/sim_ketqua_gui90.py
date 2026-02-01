# Simulate pressing the 'Kết quả' button for match 90 in the GUI
import sys, os, time
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# Prepare fake requests module to supply vMix XML
class FakeResponse:
    def __init__(self, text):
        self.text = text

class FakeRequests:
    def get(self, url, timeout=2):
        # return vMix XML with input number 1 containing required text fields
        xml = '''<vmix>
  <inputs>
    <input number="1">
      <text name="DiemA.Text">17</text>
      <text name="DiemB.Text">23</text>
      <text name="Lco.Text">3</text>
      <text name="HR1A.Text">5</text>
      <text name="HR2A.Text">6</text>
      <text name="HR1B.Text">4</text>
      <text name="HR2B.Text">2</text>
      <text name="AvgA.Text">5.667</text>
      <text name="AvgB.Text">7.667</text>
      <text name="TenA">CHÍ LONG (Chí Long)</text>
      <text name="TenB">THÀNH KHÁNH (Thanh Tự)</text>
    </input>
  </inputs>
</vmix>'''
        return FakeResponse(xml)

# Insert fake requests into sys.modules so `import requests` returns it
import types
fake_requests_module = types.ModuleType('requests')
fake_requests_module.get = FakeRequests().get
sys.modules['requests'] = fake_requests_module

# Monkeypatch GSheetClient used by the GUI to avoid real Google calls
class DummyGSheetClient:
    def __init__(self, spreadsheet_id, creds_path):
        self.spreadsheet_id = spreadsheet_id
        self.creds_path = creds_path
    def read_table(self, range_name):
        # return a list of dicts; header keys include Trận and result columns
        headers = ['Trận', 'BÀN', 'Tên VĐV A', 'Tên VĐV B', 'Điểm A', 'Điểm B', 'Lượt cơ', 'HR1A', 'HR2A', 'HR1B', 'HR2B', 'AVGA', 'AVGB']
        # create 200 rows with Trận numbers as strings
        rows = []
        for i in range(1, 201):
            d = {h: '' for h in headers}
            d['Trận'] = str(i)
            rows.append(d)
        return rows
    def write_table(self, range_name, values):
        print(f"DummyGSheetClient.write_table called: {range_name} -> {values}")
        return {'range': range_name}
    def batch_update(self, batch):
        print(f"DummyGSheetClient.batch_update called with {len(batch)} items")
        return {'responses': [{'updatedRange': batch[0].get('range') if batch else ''}]} 
    def batch_get(self, range_name):
        return [['DUMMY']]

# Now import the GUI module and replace GSheetClient
import src.gui_fullscreen_match as gui_mod
gui_mod.GSheetClient = DummyGSheetClient

# Create app
app = gui_mod.FullScreenMatchGUI()
# Configure app minimal state
sheet_id = '1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c'
app.url_var.set(f'https://docs.google.com/spreadsheets/d/{sheet_id}/edit')
app.creds_path = 'credentials.json'

# Ensure there is at least one row in match_rows
if not getattr(app, 'match_rows', None):
    print('No match_rows present; abort')
    sys.exit(2)
# Put match number 90 into first row
first_row = app.match_rows[0]
# first_row[0] is entry for Trận
try:
    first_row[0].config(state='normal')
    first_row[0].delete(0, 'end')
    first_row[0].insert(0, '90')
except Exception:
    pass
# set vmix url in column 5
try:
    first_row[5].config(state='normal')
    first_row[5].delete(0, 'end')
    first_row[5].insert(0, 'http://fake-vmix')
except Exception:
    pass

# Find the Kết quả button (index 6)
btn = first_row[6]
print('Invoking Kết quả button for match 90...')
# invoke button command
btn.invoke()
# process events a little
app.update()
# wait briefly for messagebox to be shown (modal)
time.sleep(1)
# Close app
app.destroy()
print('Done')
