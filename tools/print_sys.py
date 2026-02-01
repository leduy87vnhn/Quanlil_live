import sys, os, json
print('CWD=', os.getcwd())
print('PATH0=', sys.path[0])
print(json.dumps(sys.path, indent=2, ensure_ascii=False))
