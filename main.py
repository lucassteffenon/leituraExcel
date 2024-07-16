import os
import pandas as pd
from displayfunction import display

'''
lista_dataFrames = []
lista_nome = []
diretorio = "/Users/lucassteffenon/projetoPlanilhas"

for arq in os.listdir(diretorio):
    if arq.endswith(".xlsx"):
        nome_arquivo = "/" + arq
        caminho_completo = diretorio + nome_arquivo
        try:
            # Ler todas as folhas (sheets) do arquivo Excel
            nome_dfs = pd.read_excel(caminho_completo, sheet_name=[0, 1], engine='openpyxl')
            print("Nome do arquivo: ", arq)
            # Iterar sobre os DataFrames no dicionário e adicioná-los à lista
            for sheet_name, df in nome_dfs.items():
                lista_dataFrames.append(df)
        except Exception as e:
            print(f"Erro ao abrir o arquivo {caminho_completo}: {e}")

# Concatenar todos os DataFrames na lista
if lista_dataFrames:
    df_concatenado = pd.concat(lista_dataFrames, ignore_index=True)

# Exibir os DataFrames usando a função display
display(lista_dataFrames)
'''


lista_dataFrames = []
lista_nome = []
diretorio = "/Users/lucassteffenon/projetoPlanilhas/LANÇTOS FOLHA MATRIZ 1 PROD.xlsx"

df = pd.read_excel(diretorio, sheet_name=[0, 1], engine='openpyxl')

display(df)