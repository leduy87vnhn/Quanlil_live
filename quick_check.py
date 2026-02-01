from src.gui_fullscreen_match import FullScreenMatchGUI
app = FullScreenMatchGUI()
print('has preview_all_btn?', hasattr(app,'preview_all_btn'))
print('attrs_sample', [k for k in dir(app) if not k.startswith('__')][:60])
try:
    app._on_close_and_save_state()
except Exception:
    pass
print('done')
