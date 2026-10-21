import traceback

fname = r'C:\Users\anny9\conciliador-app\backend\integrador.py'
try:
    with open(fname, 'r', encoding='utf-8') as f:
        src = f.read()
    compile(src, fname, 'exec')
    print('ok')
except Exception:
    traceback.print_exc()
