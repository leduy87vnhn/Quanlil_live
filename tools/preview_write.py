import sys, os
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)
from src.gui_fullscreen_match import FullScreenMatchGUI

app = FullScreenMatchGUI()
batch, updates, summary = app.preview_write_all_vmix_to_sheet()
print('Summary:', summary)
print('Batch items:', len(batch))
for i, it in enumerate(batch[:10]):
    print(i+1, it.get('range'), it.get('values'))
app._on_close_and_save_state()
