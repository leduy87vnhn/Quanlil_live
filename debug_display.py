from src.gui_fullscreen_match import FullScreenMatchGUI
import time

app = FullScreenMatchGUI()
print('Children keys:', list(app.children.keys()))
# Inspect some known attributes
for attr in ['table_frame','bottom_frame','table_canvas','preview_all_btn','status_label']:
    print(attr, 'exists?', hasattr(app, attr))
# Print widget types
for k,v in app.children.items():
    print('child', k, type(v))

# Print geometry and bg where possible
try:
    print('Width x Height:', app.winfo_width(), app.winfo_height())
    print('Screen:', app.winfo_screenwidth(), app.winfo_screenheight())
except Exception as e:
    print('winfo error', e)

# keep running for 8 seconds
app.after(8000, lambda: app._on_close_and_save_state())
app.mainloop()
print('Exited mainloop')
