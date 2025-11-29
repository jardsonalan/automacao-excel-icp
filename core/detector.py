import re
import pandas as pd
from tkinter import messagebox, simpledialog
from .utils import normalize_str

def detectar_coluna_id(df):
  chaves = ['id']

  norm_cols = {col: normalize_str(col) for col in df.columns}

  for col, ncol in norm_cols.items():
    if any(chave in ncol for chave in chaves):
      return col

  for col, ncol in norm_cols.items():
    if re.search(r'\bid\b', ncol):
      return col

  cols_text = "\n".join(f"- {c}" for c in df.columns)
  messagebox.showinfo("Escolha de Coluna", "Colunas:\n\n" + cols_text)
  escolha = simpledialog.askstring("Coluna de ID", "Nome exato da coluna de ID:")

  if escolha not in df.columns:
    raise Exception(f"A coluna '{escolha}' não existe.")

  return escolha


def detectar_header_automatico(arquivo):
  df_raw = pd.read_excel(arquivo, header=None)

  for index, row in df_raw.iterrows():
    if "id" in " ".join(str(x).lower() for x in row):
      return index

  raise Exception("Nenhuma linha contendo 'ID' foi encontrada.")