import pickle
import os

STATE_PATH = os.path.join(os.path.dirname(__file__), 'ui_state.pkl')

def test_save_and_restore():
    # 1. Tạo dữ liệu giả lập
    fake_state = {
        'tengiai': 'GIẢI TEST',
        'thoigian': '01/01/2026',
        'diadiem': 'TEST LOCATION',
        'chuchay': 'CHẠY CHỮ TEST',
        'diemso': '99-88',
        'sheet_url': 'https://test',
        'creds_path': 'test_creds.json',
        'ban': 4,
        'table': [
            ['1', '1', 'A1', 'B1', '10-8', 'http://127.0.0.1:8088'],
            ['2', '2', 'A2', 'B2', '12-10', 'http://127.0.0.1:8089'],
            ['3', '3', 'A3', 'B3', '15-13', 'http://127.0.0.1:8090'],
            ['4', '4', 'A4', 'B4', '20-18', 'http://127.0.0.1:8091'],
        ],
        'preview': [
            {'type': 'vmix', 'value': 0, 'image_mode': 'fit'},
            {'type': 'vmix', 'value': 1, 'image_mode': 'fit'},
            {'type': 'image', 'value': 'test.png', 'image_mode': 'center'},
            None, None, None, None, None, None
        ]
    }
    # 2. Lưu trạng thái
    with open(STATE_PATH, 'wb') as f:
        pickle.dump(fake_state, f)
    # 3. Đọc lại trạng thái
    with open(STATE_PATH, 'rb') as f:
        loaded = pickle.load(f)
    # 4. So sánh
    assert loaded['ban'] == fake_state['ban'], f"ban: {loaded['ban']} != {fake_state['ban']}"
    assert loaded['table'] == fake_state['table'], f"table: {loaded['table']} != {fake_state['table']}"
    assert loaded['tengiai'] == fake_state['tengiai']
    assert loaded['diemso'] == fake_state['diemso']
    assert loaded['preview'][0]['type'] == 'vmix'
    print('TEST PASSED: Save/restore UI state works as expected.')

if __name__ == '__main__':
    test_save_and_restore()
