import tkinter as tk
import tempfile, os, time
try:
    from PIL import Image, ImageTk, ImageGrab
except Exception:
    Image = None
    ImageTk = None
    ImageGrab = None

# prepare sample images
tmpdir = tempfile.gettempdir()
paths = []
for i, color in enumerate([(255,0,0),(0,255,0),(0,0,255)]):
    p = os.path.join(tmpdir, f'test_preview_{i}.png')
    try:
        if Image:
            img = Image.new('RGB', (800,600), color)
            img.save(p)
        else:
            open(p, 'wb').write(b'')
    except Exception:
        open(p, 'wb').write(b'')
    paths.append(p)

root = tk.Tk()
root.title('Test Preview 3x3')
try:
    root.attributes('-fullscreen', True)
except Exception:
    try:
        root.state('zoomed')
    except Exception:
        pass

cells = []
for r in range(3):
    root.grid_rowconfigure(r, weight=1)
    for c in range(3):
        root.grid_columnconfigure(c, weight=1)
        f = tk.Frame(root, bg='#000', bd=4, relief='ridge')
        f.grid(row=r, column=c, sticky='nsew', padx=4, pady=4)
        cells.append(f)

# render images into first three cells with fit/center/cover
modes = ['fit','center','cover']
for i in range(3):
    f = cells[i]
    p = paths[i]
    try:
        img = Image.open(p)
        fw = max(10, f.winfo_screenwidth()//3)
        fh = max(10, f.winfo_screenheight()//3)
        # wait for actual geometry
        root.update_idletasks()
        fw = max(10, f.winfo_width())
        fh = max(10, f.winfo_height())
        mode = modes[i]
        if mode == 'fit':
            img2 = img.copy()
            img2.thumbnail((fw-20, fh-20), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
            tkimg = ImageTk.PhotoImage(img2)
            lbl = tk.Label(f, image=tkimg, bg='#000')
            lbl.image = tkimg
            lbl.place(relx=0.5, rely=0.5, anchor='center')
        elif mode == 'center':
            iw, ih = img.size
            scale = min(1.0, float(fw-20)/iw, float(fh-20)/ih)
            if scale < 1.0:
                img2 = img.resize((int(iw*scale), int(ih*scale)), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
            else:
                img2 = img
            tkimg = ImageTk.PhotoImage(img2)
            lbl = tk.Label(f, image=tkimg, bg='#000')
            lbl.image = tkimg
            lbl.place(relx=0.5, rely=0.5, anchor='center')
        else:
            iw, ih = img.size
            scale = max(float(fw-20)/iw, float(fh-20)/ih)
            nw = int(iw*scale)
            nh = int(ih*scale)
            img2 = img.resize((nw, nh), Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS)
            left = max(0,(nw - (fw-20))//2)
            top = max(0,(nh - (fh-20))//2)
            img3 = img2.crop((left, top, left + (fw-20), top + (fh-20)))
            tkimg = ImageTk.PhotoImage(img3)
            lbl = tk.Label(f, image=tkimg, bg='#000')
            lbl.image = tkimg
            lbl.place(relx=0.5, rely=0.5, anchor='center')
    except Exception as ex:
        tk.Label(f, text=f'Error: {ex}', fg='white', bg='red').pack(expand=True)

root.update()
# take screenshot
out = os.path.join(tmpdir, 'preview_test_capture.png')
try:
    if ImageGrab:
        x = root.winfo_rootx()
        y = root.winfo_rooty()
        w = root.winfo_width()
        h = root.winfo_height()
        img = ImageGrab.grab((x,y,x+w,y+h))
        img.save(out)
        print('Saved screenshot to', out)
    else:
        print('Pillow ImageGrab not available; no screenshot saved')
except Exception as e:
    print('Screenshot failed:', e)

root.after(1500, root.destroy)
root.mainloop()
