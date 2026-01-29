import unittest
from src.mapper import map_row_to_commands

class TestMapper(unittest.TestCase):
    def test_map_row_to_commands_basic(self):
        row = {"Name": "Nguyen Van A", "Score": "9.5", "Extra": "x"}
        field_map = {"Name": "SelectedName", "Score": "Value"}
        cmds = map_row_to_commands(row, field_map)
        self.assertIn(("SelectedName", "Nguyen Van A"), cmds)
        self.assertIn(("Value", "9.5"), cmds)

    def test_map_row_skips_empty(self):
        row = {"Name": "   ", "Score": ""}
        field_map = {"Name": "SelectedName", "Score": "Value"}
        cmds = map_row_to_commands(row, field_map)
        self.assertEqual(cmds, [])

if __name__ == "__main__":
    unittest.main()
