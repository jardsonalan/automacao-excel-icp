from .detector import detectar_coluna_id
from .extractor import extrair_id_central

def separar_por_tabela(df):
  coluna_id = detectar_coluna_id(df)

  df['ID_EXTRAIDO'] = df[coluna_id].apply(extrair_id_central)

  df_validos = df.dropna(subset=['ID_EXTRAIDO'])

  tabelas = {
    numero: df_validos[df_validos['ID_EXTRAIDO'] == numero].copy()
    for numero in sorted(df_validos['ID_EXTRAIDO'].unique())
  }

  return tabelas