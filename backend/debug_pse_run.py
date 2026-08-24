from pathlib import Path
from pse_conciliador import PseConciliador

pse_path = Path('tmp/PSE_PRUEBA.xlsx')
cruces_path = Path('tmp/CRUCES_PRUEBA.xlsx')

pse_bytes = pse_path.read_bytes()
cruces_bytes = cruces_path.read_bytes()

pc = PseConciliador(pse_bytes, cruces_bytes)
res = pc.procesar()

print('pse_entries:', len(pc.pse_entries))
print('cruces_entries:', len(pc.cruces_entries))
print('dataset:', len(res.get('dataset', [])))
print('dataset_cruces:', len(res.get('dataset_cruces', [])))
print('logs count:', len(res.get('logs', [])))
print('alertas:', res.get('alertas'))
print('resumen:', res.get('resumen'))

for row in res.get('dataset_cruces', []):
    print(row)
