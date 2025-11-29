from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def salvar_excel(tabelas, arquivo_saida):
  wb = Workbook()
  wb.remove(wb.active)

  for numero, tabela_df in tabelas.items():
    tabela_df = tabela_df.copy()

    for col in tabela_df.select_dtypes(include=['int', 'float']).columns:
      tabela_df[col] = tabela_df[col].round(2)

    ws = wb.create_sheet(title=f"Tabela_{numero}")

    for r in dataframe_to_rows(tabela_df, index=False, header=True):
      ws.append(r)

  wb.save(arquivo_saida)