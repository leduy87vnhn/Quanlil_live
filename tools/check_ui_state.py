import pickle, sys
p = r'c:\\Quanlil_live\\src\\ui_state.pkl'
try:
    with open(p,'rb') as f:
        s = pickle.load(f)
    print('ban=', s.get('ban'))
    pv = s.get('preview')
    if pv and len(pv)>0:
        print('preview0=', pv[0])
    else:
        print('preview0=', None)
except Exception as e:
    print('ERROR', e)
    sys.exit(2)
