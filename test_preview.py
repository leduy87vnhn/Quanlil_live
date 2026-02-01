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
for i, color in enumerate([(255,0,0),(0,255,0),(0,0,255)]):
    p = os.path.join(tmpdir, f'test_preview_{i}.png')
    try:
        if Image:
            img = Image.new('RGB', (800,600), color)
            img.save(p)
        else:
            # fallback: create a tiny BMP via tkinter PhotoImage? Skip if no PIL
            open(p, 'wb').write(b'')
    except Exception:
        open(p, 'wb').write(b'')
    paths.append(p)

app = FullScreenMatchGUI()
# open preview
try:
    print('app has open_preview_all?', hasattr(app, 'open_preview_all'))
    app.open_preview_all()
    preview = getattr(app, '_preview_window', None)
    if preview is None:
        print('No preview window')
    else:
        # assign images to first three cells with different modes
        preview.cell_meta[0] = {'type': 'image', 'value': paths[0], 'image_mode': 'fit'}
        preview.cell_meta[1] = {'type': 'image', 'value': paths[1], 'image_mode': 'center'}
        preview.cell_meta[2] = {'type': 'image', 'value': paths[2], 'image_mode': 'cover'}
        preview.update_idletasks()
        # trigger configure events to force render
        for f in preview.cells[:3]:
            try:
                f.event_generate('<Configure>')
            except Exception:
                pass
        # wait a moment for render
        time.sleep(0.5)
        # take screenshot if possible
        if ImageGrab and preview.winfo_ismapped():
            x = preview.winfo_rootx()
            y = preview.winfo_rooty()
            w = preview.winfo_width()
            h = preview.winfo_height()
            bbox = (x, y, x+w, y+h)
            try:
                img = ImageGrab.grab(bbox)
                out = os.path.join(tmpdir, 'preview_capture.png')
                img.save(out)
                print('Saved screenshot to', out)
            except Exception as e:
                print('Screenshot failed:', e)
        else:
            print('ImageGrab not available or preview not mapped; cannot capture screenshot')
except Exception as e:
    print('Error during preview setup:', e)

app.after(2000, app._on_close_and_save_state)
app.mainloop()
