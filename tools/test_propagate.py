import sys, os, pickle
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)
from src.gui_fullscreen_match import FullScreenMatchGUI
pkl = os.path.join(root, 'src', 'ui_state.pkl')
print('ui_state before:', pkl, os.path.exists(pkl))
with open(pkl, 'rb') as f:
    s = pickle.load(f)
print('table rows before:', len(s.get('table', [])))
# instantiate GUI to restore
app = FullScreenMatchGUI()
# call propagate_master_score
val = '9.5'
print('propagating master score ->', val)
app.propagate_master_score(val)
# save state
app._on_close_and_save_state()
# read back
with open(pkl, 'rb') as f:
    s2 = pickle.load(f)
print('table rows after:', len(s2.get('table', [])))
# print first 6 columns of first 5 rows
for i, r in enumerate(s2.get('table', [])[:10]):
    print(i, r)
