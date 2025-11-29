import re

def extrair_id_central(id_str):
  """
  Extrai o número entre E* e SA
  Ex: E7_42201_SA -> 42201
  """
  id_str = str(id_str).strip()
  match = re.search(r'_(\d+)_SA', id_str, re.IGNORECASE)
  if match:
      return int(match.group(1))
  return None