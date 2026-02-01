import pickle
import os

path = os.path.join(os.path.dirname(__file__), '..', 'src', 'ui_state.pkl')
path = os.path.abspath(path)
print('Checking:', path)
if not os.path.exists(path):
    print('File does not exist!')
    exit(1)
with open(path, 'rb') as f:
    try:
        state = pickle.load(f)
    except Exception as e:
        print('Error loading pickle:', e)
        exit(2)
print('Keys:', list(state.keys()))
for k, v in state.items():
    print(f'--- {k} ---')
    if isinstance(v, list):
        print(f'  List of {len(v)} items')
        if len(v) > 0 and isinstance(v[0], list):
            for i, row in enumerate(v):
                print(f'    Row {i}:', row)
                if i >= 2:
                    print('    ...')
                    break
        else:
            print('   ', v[:3], '...')
    else:
        print('   ', v)
