import threading, http.server, socketserver, time, os, tempfile
from src.gui_fullscreen_match import FullScreenMatchGUI

PORT = 9000
state = {'score': 0}

class VMixHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith('/API'):
            s = state['score']
            xml = f"""<vmix>
  <input number=\"1\">
    <text name=\"TenA.Text\">AAA</text>
    <text name=\"TenB.Text\">BBB</text>
    <text name=\"Tran.Text\">1</text>
    <text name=\"DiemA.Text\">{s}</text>
    <text name=\"DiemB.Text\">{s+1}</text>
    <text name=\"AVGA\">1.23</text>
    <text name=\"AVGB\">2.34</text>
    <text name=\"Lco.Text\">5</text>
  </input>
</vmix>"""
            self.send_response(200)
            self.send_header('Content-type', 'application/xml')
            self.end_headers()
            self.wfile.write(xml.encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()
    def log_message(self, format, *args):
        return

def run_server():
    with socketserver.TCPServer(('127.0.0.1', PORT), VMixHandler) as httpd:
        httpd.timeout = 1
        while not getattr(httpd, 'stopped', False):
            httpd.handle_request()

# start server
srv = threading.Thread(target=run_server, daemon=True)
srv.start()

# increment score periodically
def tick():
    for i in range(6):
        state['score'] += 1
        time.sleep(1)

tk_thread = threading.Thread(target=tick, daemon=True)
tk_thread.start()

# create GUI and fake match_rows with a vMix URL in column 5
class Dummy:
    def __init__(self, v):
        self._v = v
    def get(self):
        return self._v

app = FullScreenMatchGUI()
url = f'http://127.0.0.1:{PORT}'
# create minimal match_rows structure: [Trận, Bàn, TênA, TênB, Điểm, vMix]
app.match_rows = [[None, None, None, None, None, Dummy(url)]]
# open preview and set cell0 from row 0 and cell1 from URL directly
p = app.open_preview_all()
if not p:
    print('Không mở được preview')
else:
    app.preview_set_cell(0, 'vmix', 0)
    app.preview_set_cell(1, 'vmix', url)
    # let it run for a few seconds to observe updates
    time.sleep(4)
    # attempt screenshot if pillow available
    try:
        from PIL import ImageGrab
        if p.winfo_ismapped():
            x = p.winfo_rootx(); y = p.winfo_rooty(); w = p.winfo_width(); h = p.winfo_height()
            img = ImageGrab.grab((x,y,x+w,y+h))
            out = os.path.join(tempfile.gettempdir(), 'vmix_preview_capture.png')
            img.save(out)
            print('Saved screenshot to', out)
    except Exception as e:
        print('No screenshot:', e)

app.after(1000, app._on_close_and_save_state)
app.mainloop()
