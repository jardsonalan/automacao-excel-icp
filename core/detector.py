import re
import pandas as pd
from tkinter import messagebox, simpledialog

def detectar_coluna_id(df):
  # Encontra colunas cujo nome é exatamente "ID"
  for col in df.columns:
    if str(col).strip().upper() == "ID":
      return col

  # Se não encontrou, mostra as colunas e pede manualmente
  cols_text = "\n".join(f"- {c}" for c in df.columns)
  messagebox.showinfo("Escolha de Coluna", "Colunas encontradas:\n\n" + cols_text)
  escolha = simpledialog.askstring("Coluna de ID", "Nome exato da coluna de ID:")

  if escolha not in df.columns:
    raise Exception(f"A coluna '{escolha}' não existe.")

  return escolha

def detectar_header_automatico(arquivo):
  df_raw = pd.read_excel(arquivo, header=None)

  for index, row in df_raw.iterrows():
    # verifica se existe uma célula exatamente equal ID (desconsiderando espaços e acentos)
    for cell in row:
      if str(cell).strip().upper() == "ID":
        return index

  raise Exception("Nenhuma linha contendo a coluna 'ID' foi encontrada.")