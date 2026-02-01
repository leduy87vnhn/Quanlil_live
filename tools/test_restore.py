import pickle, os, time
p = r'c:\\Quanlil_live\\src\\ui_state.pkl'
preview = [{'type':'row_field','value':(0,'name_a'),'image_mode':'fit'}] + [None]*8
state = {
    'tengiai':'TEST',
    'thoigian':'',
    'diadiem':'',
    'chuchay':'',
    'diemso':'',
    'sheet_url':'',
    'creds_path':'',
    'ban':6,
    'table':[{'match':'T'+str(i+1)} for i in range(6)],
    'preview': preview
}
with open(p, 'wb') as f:
    pickle.dump(state, f)
print('WROTE', p)
with open(p, 'rb') as f:
    s1 = pickle.load(f)
print('BEFORE: ban=', s1.get('ban'), 'preview0=', s1.get('preview')[0])
for i in range(8,0,-1):
    print(f'waiting {i}s...')
    time.sleep(1)
with open(p, 'rb') as f:
    s2 = pickle.load(f)
print('AFTER: ban=', s2.get('ban'), 'preview0=', s2.get('preview')[0])
log = r'c:\\Quanlil_live\\src\\vmix_debug.log'
print('\n--- vmix_debug.log tail ---')
if os.path.exists(log):
    with open(log, 'r', encoding='utf-8') as lf:
        lines = lf.readlines()[-50:]
    for L in lines:
        print(L.rstrip())
else:
    print('no log')
