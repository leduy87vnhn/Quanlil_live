import os, zipfile, glob

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DIST = os.path.join(ROOT, 'dist')
RELEASES = os.path.join(ROOT, 'releases')
os.makedirs(RELEASES, exist_ok=True)

exe_candidates = glob.glob(os.path.join(DIST, '*.exe'))
exe_path = None
for e in exe_candidates:
    if os.path.basename(e).lower().startswith('quanlil_live'):
        exe_path = e
        break
if not exe_path and exe_candidates:
    exe_path = exe_candidates[0]

# try to find vmix_input1_temp.csv (project root or dist)
vmix_candidates = [os.path.join(ROOT, 'vmix_input1_temp.csv'), os.path.join(DIST, 'vmix_input1_temp.csv')]
vmix_path = None
for p in vmix_candidates:
    if os.path.exists(p):
        vmix_path = p
        break

zipname = os.path.join(RELEASES, 'Quanlil_live_v1.zip')
with zipfile.ZipFile(zipname, 'w', compression=zipfile.ZIP_DEFLATED) as z:
    if exe_path:
        z.write(exe_path, arcname=os.path.basename(exe_path))
    if vmix_path:
        z.write(vmix_path, arcname=os.path.basename(vmix_path))
    # include README
    readme = os.path.join(ROOT, 'README.md')
    if os.path.exists(readme):
        z.write(readme, arcname='README.md')

print('Created release:', zipname)
