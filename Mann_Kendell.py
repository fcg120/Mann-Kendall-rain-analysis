import pandas as pd
import numpy as np
from scipy.stats import mannwhitneyu
from tkinter import Tk, filedialog, messagebox
import tkinter as tk

# Função para carregar a planilha a partir do arquivo fornecido
def carregar_planilha_arquivo(filepath):
    dados = pd.read_excel(filepath, decimal=',')
    dados['Data'] = pd.to_datetime(dados['Data'], format='%d/%m/%Y')
    if 'Data' not in dados.columns or 'Precipitação' not in dados.columns:
        raise ValueError("A planilha deve conter as colunas 'Data' e 'Precipitação'")
    return dados

# Função para realizar o teste de Mann-Kendall
def mann_kendall_test(data):
    n = len(data)
    s = 0
    for k in range(n - 1):
        for j in range(k + 1, n):
            s += np.sign(data[j] - data[k])
    unique_data = np.unique(data)
    g = len(unique_data)
    if n == g:  # Nenhum dado duplicado
        var_s = (n * (n - 1) * (2 * n + 5)) / 18
    else:  # Duplicatas presentes
        tp = np.bincount(np.searchsorted(unique_data, data))
        var_s = (n * (n - 1) * (2 * n + 5) - np.sum(tp * (tp - 1) * (2 * tp + 5))) / 18
    if s > 0:
        z = (s - 1) / np.sqrt(var_s)
    elif s < 0:
        z = (s + 1) / np.sqrt(var_s)
    else:
        z = 0
    p = 2 * (1 - mannwhitneyu(np.arange(n), data).pvalue)
    h = abs(z) > 1.96
    return h, p, z

# Função para exibir os resultados em uma janela de diálogo
def exibir_resultados(filepath):
    try:
        dados = carregar_planilha_arquivo(filepath)
        if dados is not None:
            h, p, z = mann_kendall_test(dados['Precipitação'].values)
            resultado = f'Teste Mann-Kendall:\n  Tendência significativa: {"Sim" if h else "Não"}\n  Valor p: {p}\n  Estatística Z: {z}\n'
            messagebox.showinfo("Resultados do Teste Mann-Kendall", resultado)
        else:
            messagebox.showerror("Erro", "Nenhum dado carregado.")
    except Exception as e:
        messagebox.showerror("Erro ao carregar os dados", str(e))

# Função para abrir o diálogo de seleção de arquivo e processar a planilha
def carregar_arquivo():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if file_path:
        exibir_resultados(file_path)

# Interface gráfica principal
def main():
    root = tk.Tk()
    root.title("Análise de Tendências de Precipitação")
    
    # Adicionar botão para carregar planilha
    button_carregar = tk.Button(root, text="Carregar Planilha", command=carregar_arquivo)
    button_carregar.pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    main()
