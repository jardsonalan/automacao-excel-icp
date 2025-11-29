import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

from core.detector import detectar_header_automatico
from core.processor import separar_por_tabela
from core.exporter import salvar_excel

def executar_processamento(entrada_arquivo):
  try:
    arquivo = entrada_arquivo.get()
    if not arquivo:
      messagebox.showerror("Erro", "Selecione um arquivo Excel.")
      return

    linha_header = detectar_header_automatico(arquivo)
    df = pd.read_excel(arquivo, header=linha_header)

    tabelas = separar_por_tabela(df)

    if not tabelas:
      messagebox.showinfo("Resultado", "Nenhum ID válido encontrado.")
      return

    arquivo_saida = arquivo.replace(".xlsx", "_separado.xlsx")
    salvar_excel(tabelas, arquivo_saida)

    messagebox.showinfo("Sucesso", f"Arquivo gerado:\n{arquivo_saida}")

  except Exception as e:
    messagebox.showerror("Erro", str(e))


def iniciar_interface():
  janela = tk.Tk()
  janela.title("Separador de IDs")
  janela.geometry("500x240")

  tk.Label(janela, text="Arquivo Excel:").pack(pady=(10, 0))
  entrada_arquivo = tk.Entry(janela, width=60)
  entrada_arquivo.pack(padx=10)

  def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
      title="Selecione a planilha Excel",
      filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    entrada_arquivo.delete(0, tk.END)
    entrada_arquivo.insert(0, caminho)

  tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo).pack(pady=8)
  tk.Button(
    janela,
    text="Separar IDs",
    command=lambda: executar_processamento(entrada_arquivo),
    bg="green",
    fg="white"
  ).pack(pady=12)

  janela.mainloop()