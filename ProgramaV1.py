import pandas as pd
import pyodbc
import tkinter as tk
from tkinter import ttk

# Realizar conexão com banco de dados SQL
dados_conexao = (
    "Driver={SQL SERVER};"
    "Server=LucasMachine;"
    "Database=ALGAR_RELATORIO_GERAL;"
)

def puxa_banco():
    conexao = pyodbc.connect(dados_conexao)
    print("Conexão bem sucedida")
    cursor = conexao.cursor()

    nome_site = entry_descricao.get()

    consulta = f"""SELECT * FROM  [dbo].['DADOS DETENTORAS$']
    WHERE ESTACAO_SEM_CODIGO = '{nome_site}'"""

    cursor.execute(consulta)

    linhas = cursor.fetchall()

    print(linhas)

    for linha in linhas:
        endereco_site = [linha[5]]

    for linha in linhas:
        print("A DETENTORA DO SITE {} é: {}".format(nome_site,linha[1]))

    tabela = pd.read_excel(r"C:\Users\Lucas\OneDrive\Área de Trabalho\Projetos Algar\Modelos FSCIs\FSCI-Modelo-ATC.xlsx")
    tabela.loc[8,10] = endereco_site
    print(tabela)

    tabela.to_excel("PlanilhaNova2.xlsx", index=False)

janela = tk.Tk()

janela.title('Consulta Detentora')

label_nomesite = tk.Label(text="Nome do site")
label_nomesite.grid(row=1, column=0,padx = 10, pady=10, sticky='nswe', columnspan =4 )

entry_descricao = tk.Entry()
entry_descricao.grid(row=2,column=0, padx=10, pady=10, sticky='nswe', columnspan = 4)

botao_realizar_consulta = tk.Button(text="Realizar Consulta", command=puxa_banco)
botao_realizar_consulta.grid(row=5,column=0,padx = 10, pady=10,sticky='nswe', columnspan =4)

janela.mainloop()
