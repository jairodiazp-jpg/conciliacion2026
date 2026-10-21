import importlib.util

spec = importlib.util.spec_from_file_location('integrador', 'C:/Users/anny9/conciliador-app/backend/integrador.py')
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
print('integrador loaded')
