import time, os, sys
# ensure workspace root is on sys.path so `src` package is importable when running from tools/
root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root not in sys.path:
    sys.path.insert(0, root)
from src.gui_fullscreen_match import FullScreenMatchGUI

try:
    from PIL import ImageGrab
except Exception:
    ImageGrab = None

app = FullScreenMatchGUI()

def do_capture():
    try:
        app.update()
        app.update_idletasks()
        # wait a bit for geometry to settle
        time.sleep(0.4)
        x = app.winfo_rootx()
        y = app.winfo_rooty()
        w = app.winfo_width()
        h = app.winfo_height()
        # if geometry still tiny, try maximize then measure
        if w <= 1 or h <= 1:
            try:
                app.state('zoomed')
                app.update()
                time.sleep(0.2)
                w = app.winfo_width(); h = app.winfo_height()
            except Exception:
                pass
        bbox = (x, y, x + max(1, w), y + max(1, h))
        out = os.path.join(os.getcwd(), 'gui_screenshot.png')
        if ImageGrab:
            img = ImageGrab.grab(bbox)
            img.save(out)
            print('Saved screenshot to', out)
        else:
            print('Pillow ImageGrab not available; skipping image capture. Geometry:', bbox)
    except Exception as e:
        print('Capture error:', e)
    finally:
        try:
            app._on_close_and_save_state()
        except Exception:
            try:
                app.destroy()
            except Exception:
                pass

app.after(600, do_capture)
app.mainloop()
