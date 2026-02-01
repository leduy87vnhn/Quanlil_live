import pickle, os
p=r'c:\\Quanlil_live\\src\\ui_state.pkl'
print('path=', p)
print('exists=', os.path.exists(p))
if os.path.exists(p):
    import pickle, os
    p = r'c:\Quanlil_live\src\ui_state.pkl'
    print('path=', p)
    print('exists=', os.path.exists(p))
    if os.path.exists(p):
        try:
            with open(p, 'rb') as f:
                s = pickle.load(f)
                print('keys=', list(s.keys()))
                print('table_rows=', len(s.get('table', [])))
                print('preview=', bool(s.get('preview')))
                pr = s.get('preview')
                if pr:
                    for i, c in enumerate(pr[:9]):
                        print(i, c)
        except Exception as e:
            print('error=', e)
    else:
        print('no file')
