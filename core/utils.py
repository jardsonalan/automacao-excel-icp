import re
import unicodedata

def normalize_str(s):
  s = str(s).strip()
  s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII')
  s = re.sub(r'\s+', ' ', s).strip().lower()
  return s