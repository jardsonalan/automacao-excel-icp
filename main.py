import pandas as pd
import re
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ---- utilitário para normalizar strings (remove acentos/espacos, lower) ----
def normalize_str(s):
    s = str(s).strip()
    # remove acentos
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII')
    # remove espaços extras e deixa minusculo
    s = re.sub(r'\s+', ' ', s).strip().lower()
    return s

# ---- extrair número final ----
def extrair_numero_final(id_str):
    id_str = str(id_str).strip()
    match = re.search(r'(\d+)$', id_str)
    return int(match.group(1)) if match else None

# ---- detectar coluna com heurísticas + fallback para escolher manualmente ----
def detectar_coluna_id(df):
    # lista de palavras chaves possíveis (normalizadas)
    chaves = ['id', 'amostra', 'codigo', 'codigo_amostra', 'cod', 'código', 'sample', 'codigoamostra', 'id_amostra']
    # precompute normalized column names
    norm_cols = {col: normalize_str(col) for col in df.columns}

    # 1) procura por coluna cuja normalized contenha exatamente 'id' ou outras chaves
    for col, ncol in norm_cols.items():
        for chave in chaves:
            # usar "contains" para pegar variações como "id amostra", "id_amostra" etc.
            if chave in ncol:
                return col

    # 2) se não encontrou, tentar encontrar qualquer coluna que termine/comece com 'id'
    for col, ncol in norm_cols.items():
        if re.search(r'\bid\b', ncol):
            return col

    # 3) fallback: mostrar ao usuário a lista de colunas (original) e solicitar entrada
    cols_text = "\n".join(f"- {c}" for c in df.columns)
    msg = ("Não foi possível detectar automaticamente a coluna de ID.\n\n"
           "Colunas encontradas na planilha:\n\n" + cols_text +
           "\n\nDigite o nome exato da coluna que contém os IDs (copie/cole como aparece acima).")
    # mostra primeiro um info com as colunas (pode ser grande)
    messagebox.showinfo("Escolha de Coluna", msg)
    # pede ao usuário digitar o nome da coluna (exatamente como está no Excel)
    escolha = simpledialog.askstring("Coluna de ID", "Nome da coluna que contém o ID (exato):")
    if not escolha:
        raise Exception("Nenhuma coluna selecionada. Operação cancelada.")
    # valida escolha
    if escolha not in df.columns:
        # tentar encontrar por normalização: se a entrada do usuário foi sem acento/maiúsculas
        for col in df.columns:
            if normalize_str(col) == normalize_str(escolha):
                return col
        raise Exception(f"A coluna '{escolha}' não existe na planilha.")
    return escolha

# ---- separa conforme número final ----
def separar_por_tabela(df):
    coluna_id = detectar_coluna_id(df)

    df['NUM_FINAL'] = df[coluna_id].apply(extrair_numero_final)

    df_validos = df.dropna(subset=['NUM_FINAL'])

    tabelas = {
        numero: df_validos[df_validos['NUM_FINAL'] == numero].copy()
        for numero in sorted(df_validos['NUM_FINAL'].unique())
    }

    return tabelas

# ---- salva excel com abas ----
def salvar_excel(tabelas, arquivo_saida):
    wb = Workbook()
    wb.remove(wb.active)

    for numero, tabela_df in tabelas.items():
        # ---- FORMATAÇÃO: arredondar todos os números para 2 casas ----
        tabela_df = tabela_df.copy()
        for col in tabela_df.select_dtypes(include=['float', 'int']).columns:
            tabela_df[col] = tabela_df[col].round(2)

        ws = wb.create_sheet(title=f"Tabela_{numero}")

        for r in dataframe_to_rows(tabela_df, index=False, header=True):
            ws.append(r)

    wb.save(arquivo_saida)

# ---- UI ----
def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    entrada_arquivo.delete(0, tk.END)
    entrada_arquivo.insert(0, caminho)

def detectar_header_automatico(arquivo):
    # Lê tudo sem header
    df_raw = pd.read_excel(arquivo, header=None)

    # Procura a linha que contém a palavra "ID"
    for index, row in df_raw.iterrows():
        # normaliza a linha inteira
        texto = " ".join(str(x).lower() for x in row)

        if "id" in texto:
            return index  # essa será usada como a linha do cabeçalho

    raise Exception("Nenhuma linha contendo 'ID' foi encontrada no arquivo.")


def executar():
    try:
        arquivo = entrada_arquivo.get()
        if not arquivo:
            messagebox.showerror("Erro", "Selecione um arquivo Excel.")
            return

        # 1. Detecta automaticamente a linha do cabeçalho
        linha_header = detectar_header_automatico(arquivo)

        # 2. Recarrega o DataFrame usando aquela linha como cabeçalho
        df = pd.read_excel(arquivo, header=linha_header)

        # 3. Continua o fluxo normal
        tabelas = separar_por_tabela(df)

        if not tabelas:
            messagebox.showinfo("Resultado", "Nenhum ID válido com número final encontrado.")
            return

        arquivo_saida = arquivo.replace(".xlsx", "_separado.xlsx")
        salvar_excel(tabelas, arquivo_saida)
        messagebox.showinfo("Sucesso", f"Arquivo gerado:\n{arquivo_saida}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# ---- janela ----
janela = tk.Tk()
janela.title("Separador Automático de Tabelas por ID (robusto)")
janela.geometry("500x240")

tk.Label(janela, text="Arquivo Excel:").pack(pady=(10,0))
entrada_arquivo = tk.Entry(janela, width=60)
entrada_arquivo.pack(padx=10)

tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo).pack(pady=8)
tk.Button(janela, text="Separar IDs em Tabelas", command=executar, bg="green", fg="white").pack(pady=12)

janela.mainloop()
