import ast
fname = r'C:\Users\anny9\conciliador-app\backend\integrador.py'
src = open(fname, 'r', encoding='utf-8').read()
try:
    ast.parse(src)
    print('ast ok')
except SyntaxError as e:
    print('SyntaxError', e.msg, 'line', e.lineno, 'offset', e.offset)
    print('\n'.join(f"{i+1}: {line}" for i, line in enumerate(src.splitlines())[:e.lineno+2]))
