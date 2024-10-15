import pandas as pd
import os
import re
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox

# Variáveis globais para armazenar os diretórios
diretorio_entrada = None
arquivo_saida = None


# Função para abrir o seletor de diretório
def selecionar_diretorio():
    global diretorio_entrada
    root = tk.Tk()
    root.withdraw()
    diretorio_entrada = filedialog.askdirectory(title='Selecione o diretório dos arquivos Excel')


# Função para abrir o seletor de arquivo de saída
def selecionar_arquivo_saida():
    global arquivo_saida
    root = tk.Tk()
    root.withdraw()
    arquivo_saida = filedialog.asksaveasfilename(title='Selecione o local para salvar o arquivo CSV filtrado',
                                                 defaultextension='.csv', filetypes=[("CSV files", "*.csv")])


# Função para iniciar o processo de filtragem e salvar o resultado
def iniciar_programa():
    if not diretorio_entrada or not arquivo_saida:
        messagebox.showerror("Erro", "Por favor, selecione o diretório dos arquivos Excel e o local de saída.")
        return

    try:
        # DataFrame vazio para armazenar todas as linhas que atendem ao critério
        df_combinado = pd.DataFrame()

        # Tabelas que devem ser percorridas
        tabelas_desejadas = ['LANÇTOS FOLHA', 'LANÇTOS PROVISÃO']

        # Função para remover caracteres especiais
        def remover_caracteres_especiais(texto):
            if isinstance(texto, str):
                nfkd = unicodedata.normalize('NFKD', texto)
                return "".join([c for c in nfkd if not unicodedata.combining(c)])
            return texto

        # Percorrer todos os arquivos do diretório
        for arquivo in os.listdir(diretorio_entrada):
            if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
                # Ler o arquivo Excel, incluindo todas as tabelas
                caminho_arquivo = os.path.join(diretorio_entrada, arquivo)
                df_dict = pd.read_excel(caminho_arquivo, sheet_name=None,header=None)

                # Percorrer apenas as tabelas desejadas do arquivo Excel
                for sheet_name in tabelas_desejadas:
                    if sheet_name in df_dict:
                        df = df_dict[sheet_name]
                        # Percorrer linha por linha do DataFrame
                        for index, row in df.iterrows():
                            # Verificar se a coluna B (índice 1) contém qualquer número e a coluna E (índice 4) está preenchida com um número diferente de zero
                            if pd.notna(row.iloc[1]) and pd.notna(row.iloc[4]):
                                if re.search(r'\d', str(row.iloc[1])):
                                    valor_e = pd.to_numeric(row.iloc[4], errors='coerce')
                                    if pd.notna(valor_e) and valor_e != 0:
                                        # Adicionar a linha inteira ao DataFrame combinado
                                        df_combinado = pd.concat([df_combinado, pd.DataFrame([row])], ignore_index=True)

        # Salvar o DataFrame combinado em um novo arquivo CSV
        df_combinado.to_csv(arquivo_saida, index=False, sep=';')

        # Abrir o arquivo salvo, remover colunas depois da coluna 5 e substituir caracteres especiais
        df_resultado = pd.read_csv(arquivo_saida, sep=';')
        df_resultado = df_resultado.iloc[:, :5]

        # Aplicar a função de remover caracteres especiais em todas as colunas
        for coluna in df_resultado.columns:
            df_resultado[coluna] = df_resultado[coluna].apply(remover_caracteres_especiais)

        # Salvar novamente o arquivo atualizado em CSV
        df_resultado.to_csv(arquivo_saida, index=False, sep=';')

        messagebox.showinfo("Sucesso", f'O arquivo {arquivo_saida} foi atualizado com sucesso.')
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro durante a execução: {str(e)}")


# Interface gráfica principal
def criar_interface():
    root = tk.Tk()
    root.title("Filtro de Arquivos Excel")
    root.geometry("500x300")

    label = tk.Label(root, text="Bem-vindo ao Filtro de Arquivos Excel!", font=("Helvetica", 18, "bold"))
    label.pack(pady=20)

    btn_selecionar_diretorio = tk.Button(root, text="Selecionar Diretório dos Arquivos", command=selecionar_diretorio, font=("Helvetica", 12), bg="#4CAF50", fg="white", padx=10, pady=5)
    btn_selecionar_diretorio.pack(pady=10)

    btn_selecionar_arquivo_saida = tk.Button(root, text="Selecionar Arquivo de Saída", command=selecionar_arquivo_saida, font=("Helvetica", 12), bg="#2196F3", fg="white", padx=10, pady=5)
    btn_selecionar_arquivo_saida.pack(pady=10)

    btn_iniciar = tk.Button(root, text="Iniciar Programa", command=iniciar_programa, font=("Helvetica", 12, "bold"), bg="#f44336", fg="white", padx=10, pady=10)
    btn_iniciar.pack(pady=20)

    root.mainloop()

# Iniciar a interface gráfica
criar_interface()