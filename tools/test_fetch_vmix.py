import threading, time, http.server, socketserver, sys, os
# ensure workspace root on path
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)
from src.gui_fullscreen_match import FullScreenMatchGUI

VMIX_PORT = 9001
VMIX_XML = '''<?xml version="1.0" encoding="utf-8"?>
<vmix>
  <inputs>
    <input number="1" title="Input1">
      <text name="TenA">Nguyen Van A</text>
      <text name="TenB">Tran Thi B</text>
      <text name="DiemA">9</text>
      <text name="DiemB">8</text>
    </input>
  </inputs>
</vmix>
'''

class SimpleHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith('/API/') or self.path == '/API':
            self.send_response(200)
            self.send_header('Content-Type', 'text/xml')
            self.send_header('Content-Length', str(len(VMIX_XML.encode('utf-8'))))
            self.end_headers()
            self.wfile.write(VMIX_XML.encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()
    def log_message(self, format, *args):
        return

def run_server(stop_event):
    with socketserver.TCPServer(('127.0.0.1', VMIX_PORT), SimpleHandler) as httpd:
        httpd.timeout = 0.5
        while not stop_event.is_set():
            httpd.handle_request()

def main():
    stop = threading.Event()
    t = threading.Thread(target=run_server, args=(stop,), daemon=True)
    t.start()
    try:
        app = FullScreenMatchGUI()
        # ensure we have rows
        if not getattr(app, 'match_rows', None):
            print('No match_rows available')
            return 2
        # set first row vmix url to our server
        try:
            app.match_rows[0][5].delete(0, 'end')
            app.match_rows[0][5].insert(0, f'http://127.0.0.1:{VMIX_PORT}')
        except Exception as e:
            print('Failed to set vmix url:', e)
            app._on_close_and_save_state()
            return 3
        # call fetch (sync) and check
        updated = app.fetch_all_vmix_to_table()
        print('fetch_all_vmix_to_table returned', updated)
        valA = app.match_rows[0][2].get() if hasattr(app.match_rows[0][2], 'get') else None
        valB = app.match_rows[0][3].get() if hasattr(app.match_rows[0][3], 'get') else None
        print('Tên A:', valA)
        print('Tên B:', valB)
        # cleanup
        app._on_close_and_save_state()
        stop.set()
        t.join(timeout=1)
        # expected values
        ok = (updated >= 1) and (valA and 'Nguyen' in valA) and (valB and 'Tran' in valB)
        print('TEST OK' if ok else 'TEST FAIL')
        return 0 if ok else 4
    except Exception as e:
        print('Test error', e)
        stop.set()
        return 5

if __name__ == '__main__':
    sys.exit(main())
