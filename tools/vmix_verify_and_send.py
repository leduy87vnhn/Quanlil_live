import requests, xml.etree.ElementTree as ET, sys, time, os

vmix_url = 'http://CHUNG:8088'
log_path = os.path.join(os.getcwd(), 'vmix_debug.log')

def log(msg):
    s = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n"
    print(s, end='')
    try:
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(s)
    except Exception:
        pass

# GET API
try:
    r = requests.get(f"{vmix_url}/API/", timeout=5)
    log(f"GET {vmix_url}/API/ -> {r.status_code}")
    root = ET.fromstring(r.text)
    input1 = root.find(".//input[@number='1']")
    if input1 is None:
        log('No input[1] found in vMix API')
        sys.exit(1)
    texts = {t.attrib.get('name'): (t.text or '') for t in input1.findall('text')}
    log(f'Parsed Input 1 texts: {texts}')
except Exception as e:
    log(f'Error fetching vMix API: {e}')
    sys.exit(1)

# Prepare mapping and test values (use current values as base)
mapping = {
    'TenA.Text': texts.get('TenA.Text', 'TestA'),
    'TenB.Text': texts.get('TenB.Text', 'TestB'),
    'DiemA.Text': texts.get('DiemA.Text', '0'),
    'DiemB.Text': texts.get('DiemB.Text', '0'),
    'Lco.Text': texts.get('Lco.Text', '1'),
    'HR1A.Text': texts.get('HR1A.Text', ''),
    'HR2A.Text': texts.get('HR2A.Text', ''),
    'HR1B.Text': texts.get('HR1B.Text', ''),
    'HR2B.Text': texts.get('HR2B.Text', ''),
}

# Send SetText for each mapped field and log responses
session = requests.Session()
for name, val in mapping.items():
    params = {'Function': 'SetText', 'Input': 1, 'SelectedName': name, 'Value': val}
    try:
        resp = session.get(f"{vmix_url}/API/", params=params, timeout=4)
        log(f"SET {name}='{val}' -> {resp.status_code}")
    except Exception as e:
        log(f"Error SET {name}: {e}")

log('Done')
