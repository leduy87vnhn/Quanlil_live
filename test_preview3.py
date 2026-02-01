from src.gui_fullscreen_match import FullScreenMatchGUI
import tempfile, os, time
try:
    from PIL import Image, ImageGrab
except Exception:
    Image = None
    ImageGrab = None

# create sample images
tmpdir = tempfile.gettempdir()
paths = []
for i, color in enumerate([(200,120,60),(60,180,120),(60,120,200)]):
    p = os.path.join(tmpdir, f'test_preview_auto_{i}.png')
    try:
        if Image:
            img = Image.new('RGB', (1024,768), color)
            img.save(p)
        else:
            open(p, 'wb').write(b'')
    except Exception:
        open(p, 'wb').write(b'')
    paths.append(p)

app = FullScreenMatchGUI()
# programmatic open preview
def do_preview_actions():
    p = app.preview_open()
    if not p:
        print('Không mở được preview (sẽ thử lại)')
        return
    print('Preview opened')
    app.preview_set_cell(0, 'image', paths[0], image_mode='fit')
    app.preview_set_cell(1, 'image', paths[1], image_mode='center')
    app.preview_set_cell(2, 'image', paths[2], image_mode='cover')
    app.update_idletasks()
    # try screenshot shortly after
    def try_screenshot():
        p2 = getattr(app, '_preview_window', None)
        if not p2:
            print('Preview window not found for screenshot')
            return
        if ImageGrab and p2.winfo_ismapped():
            x = p2.winfo_rootx(); y = p2.winfo_rooty(); w = p2.winfo_width(); h = p2.winfo_height()
            try:
                img = ImageGrab.grab((x,y,x+w,y+h))
                out = os.path.join(tmpdir, 'preview_auto_capture.png')
                img.save(out)
                print('Saved screenshot to', out)
            except Exception as e:
                print('Screenshot failed:', e)
        else:
            print('Không thể chụp ảnh (ImageGrab không khả dụng hoặc preview chưa hiển thị)')

    app.after(600, try_screenshot)

app.after(200, do_preview_actions)
app.after(2500, app._on_close_and_save_state)
app.mainloop()
