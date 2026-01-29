# Runner to perform a one-shot full update (non-GUI blocking)
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
from gui_fullscreen_match import FullScreenMatchGUI
import traceback, sys

print('Runner: creating FullScreenMatchGUI instance...')
try:
    app = FullScreenMatchGUI()
    # Configure Google Sheet URL and credentials (set by runner)
    SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ACFOmGQDQrAJvFWXJYIEdvhUJJFQC7LViVvt-AHjl-c/edit?gid=778091296'
    app.url_var.set(SHEET_URL)
    # Prefer credentials.json in repo root if exists
    repo_cred = os.path.join(os.path.dirname(__file__), 'credentials.json')
    if os.path.exists(repo_cred):
        app.creds_path = repo_cred
    # Reload matches to populate `sheet_rows`
    print('Runner: reloading matches from Google Sheet...')
    try:
        app.reload_matches()
    except Exception as e:
        print('reload_matches error:', e)
    # First try exporting vmix input CSV explicitly so we can inspect it
    print('Runner: calling export_vmix_input1_to_tempfile()...')
    try:
        app.export_vmix_input1_to_tempfile()
    except Exception as e:
        print('export_vmix_input1_to_tempfile error:', e)
    print('Runner: now calling update_gsheet_with_vmix_full()...')
    try:
        app.update_gsheet_with_vmix_full()
    except Exception as e:
        print('update_gsheet_with_vmix_full error:', e)
    print('Runner: update finished.')
    try:
        app.destroy()
    except Exception:
        pass
except Exception as e:
    print('Runner error:', e)
    traceback.print_exc()
    sys.exit(1)
