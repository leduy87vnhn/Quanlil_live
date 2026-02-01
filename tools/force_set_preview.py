import os, pickle, shutil
p = os.path.join(os.getcwd(), 'src', 'ui_state.pkl')
if not os.path.exists(p):
    print('no ui_state.pkl found at', p)
    raise SystemExit(1)
# backup
bak = p + '.bak'
shutil.copy2(p, bak)
print('backup created:', bak)
with open(p, 'rb') as f:
    s = pickle.load(f)
print('keys before:', list(s.keys()))
# create a sample preview: 9 cells with first cell as vmix row 0
preview = []
for i in range(9):
    if i == 0:
        preview.append({'type':'vmix','value':0,'image_mode':'fit'})
    else:
        preview.append({'type':None,'value':None,'image_mode':'fit'})
s['preview'] = preview
with open(p, 'wb') as f:
    pickle.dump(s, f)
print('wrote preview into', p)
print('keys after:', list(s.keys()))
