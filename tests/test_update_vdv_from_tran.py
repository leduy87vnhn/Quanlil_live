import unittest
from unittest.mock import MagicMock
from src.gui_fullscreen_match import FullScreenMatchGUI

class TestUpdateVdvFromTran(unittest.TestCase):
    def setUp(self):
        self.app = FullScreenMatchGUI()
        # Giả lập sheet_rows với tiêu đề cột linh hoạt
        self.app.sheet_rows = [
            {'Trận': '1', 'VĐV A': 'A1', 'VĐV B': 'B1'},
            {'Trận': '2', 'VĐV A': 'A2', 'VĐV B': 'B2'},
            {'Trận': '03', 'VĐV A': 'A3', 'VĐV B': 'B3'},
        ]
        # Giả lập match_rows: mỗi dòng là list [e_tran, e_ban, e_a, e_b, btn, status]
        class DummyEntry:
            def __init__(self, val=''):
                self.val = val
                self.state = 'normal'
            def get(self): return self.val
            def insert(self, idx, v): self.val = v
            def delete(self, a, b): self.val = ''
            def config(self, **kwargs): self.state = kwargs.get('state', self.state)
        class DummyLabel:
            def __init__(self): self.text = ''
            def config(self, **kwargs): self.text = kwargs.get('text', self.text)
        self.DummyEntry = DummyEntry
        self.DummyLabel = DummyLabel
        self.app.match_rows = [
            [DummyEntry('1'), DummyEntry('Bàn 1'), DummyEntry(), DummyEntry(), None, DummyLabel()],
            [DummyEntry('03'), DummyEntry('Bàn 2'), DummyEntry(), DummyEntry(), None, DummyLabel()],
            [DummyEntry('2'), DummyEntry('Bàn 3'), DummyEntry(), DummyEntry(), None, DummyLabel()],
            [DummyEntry('99'), DummyEntry('Bàn 4'), DummyEntry(), DummyEntry(), None, DummyLabel()],
        ]
    def test_update_vdv_from_tran(self):
        # Dòng 0: nhập '1' phải ra A1, B1
        self.app.update_vdv_from_tran(0)
        self.assertEqual(self.app.match_rows[0][2].val, 'A1')
        self.assertEqual(self.app.match_rows[0][3].val, 'B1')
        # Dòng 1: nhập '03' phải ra A3, B3
        self.app.update_vdv_from_tran(1)
        self.assertEqual(self.app.match_rows[1][2].val, 'A3')
        self.assertEqual(self.app.match_rows[1][3].val, 'B3')
        # Dòng 2: nhập '2' phải ra A2, B2
        self.app.update_vdv_from_tran(2)
        self.assertEqual(self.app.match_rows[2][2].val, 'A2')
        self.assertEqual(self.app.match_rows[2][3].val, 'B2')
        # Dòng 3: nhập '99' không có, phải rỗng và báo lỗi
        self.app.update_vdv_from_tran(3)
        self.assertEqual(self.app.match_rows[3][2].val, '')
        self.assertEqual(self.app.match_rows[3][3].val, '')
        self.assertIn('không', self.app.match_rows[3][5].text.lower())

if __name__ == '__main__':
    unittest.main()
