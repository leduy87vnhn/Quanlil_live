import urllib.request
import xml.etree.ElementTree as ET
import sys

url = 'http://CHUNG:8088/API/'
try:
    with urllib.request.urlopen(url, timeout=5) as resp:
        data = resp.read().decode('utf-8', errors='replace')
    root = ET.fromstring(data)
    input1 = root.find(".//input[@number='1']")
    def get_field(name):
        if input1 is None:
            return None
        for t in input1.findall('text'):
            if t.attrib.get('name') == name:
                return t.text if t.text is not None else ''
        return None
    v = get_field('DiemA.Text')
    if v is None:
        # fallback
        v = get_field('DiemA') or get_field('ScoreA')
    if v is None:
        print('NOT_FOUND')
    else:
        print(v)
except Exception as e:
    print('ERROR:'+str(e))
