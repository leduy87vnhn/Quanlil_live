import unittest
from src.vmix_parser import extract_fields_from_state

class TestVmixParser(unittest.TestCase):
    def test_extract_simple_tags(self):
        xml = "<vmix><SelectedName>Nguyen Van A</SelectedName><Value>9.50</Value></vmix>"
        fields = extract_fields_from_state(xml, ["SelectedName", "Value"])
        self.assertEqual(fields["SelectedName"], "Nguyen Van A")
        self.assertEqual(fields["Value"], "9.50")

    def test_extract_by_name_attribute(self):
        xml = "<vmix><field name=\"SelectedName\">Tran B</field><field name=\"Value\">8.25</field></vmix>"
        fields = extract_fields_from_state(xml, ["SelectedName", "Value"])
        self.assertEqual(fields["SelectedName"], "Tran B")
        self.assertEqual(fields["Value"], "8.25")

if __name__ == "__main__":
    unittest.main()
