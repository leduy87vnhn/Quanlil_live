import sys, os, pickle
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)
p = os.path.join(root, 'src', 'ui_state.pkl')
print('path=', p)
if not os.path.exists(p):
    print('no ui_state.pkl')
    sys.exit(1)
with open(p, 'rb') as f:
    s = pickle.load(f)
# ensure preview list length >=9
pr = s.get('preview')
if not isinstance(pr, list) or len(pr) < 9:
    pr = [{'type': None, 'value': None, 'image_mode': 'fit'} for _ in range(9)]
# set cell 0 to vmix row 0
pr[0] = {'type': 'vmix', 'value': 0, 'image_mode': 'fit'}
# set cell 4 to image path if exists
img_path = os.path.join(root, 'assets', 'logo.png')
if os.path.exists(img_path):
    pr[4] = {'type': 'image', 'value': img_path, 'image_mode': 'fit'}
s['preview'] = pr
with open(p + '.bak', 'wb') as f:
    pickle.dump(s, f)
print('backup written', p + '.bak')
with open(p, 'wb') as f:
    pickle.dump(s, f)
print('wrote preview into', p)
# instantiate GUI to trigger restore
from src.gui_fullscreen_match import FullScreenMatchGUI
print('instantiating GUI (no mainloop)')
app = FullScreenMatchGUI()
print('done')
app._on_close_and_save_state()
print('saved state after instantiation')
