import tkinter as tk
from tkinter import ttk
import datetime as dt
import pyodbc

#Realizar conexão com banco de dados SQL
dados_conexao = (
    "Driver={SQL SERVER};"
    "Server=LucasMachine;"
    "Database=CADASTRO_CLIENTES;"
)

conexao = pyodbc.connect(dados_conexao)
print("Conexão bem sucedida")
cursor = conexao.cursor()

lista_documentos = ["RG", "CPF"]
lista_clientes = []

janela = tk.Tk()

#Criação da função

def cadastrar_cliente():
    nomecliente = entry_nomecliente.get()
    tipo = combobox_selecionar_tipo.get()
    numerodocumento = entry_numerodocumento.get()
    data_criacao = dt.datetime.now()
    data_criacao = data_criacao.strftime("%d/%m/%Y %H:%M")
    # Cadastrar em uma lista (necessário depois modificar para cadastrar em um banco de dados)
    clientes = len(lista_clientes)+1
    clientes_str = "Cliente Número-{}".format(clientes)
    lista_clientes.append((clientes_str,nomecliente,tipo,numerodocumento,data_criacao))

    comando = f"""INSERT INTO [CADASTRO_CLIENTES]([Codigo do Cliente],[Nome do Cliente],[Tipo do Documento],[Numero do Documento],[Data Cadastro])
    VALUES ('{clientes}','{nomecliente}','{tipo}','{numerodocumento}','{data_criacao}')"""
    cursor.execute(comando)
    cursor.commit()

def deletar_cliente():
    nomecliente = entry_nomecliente.get()
    tipo = combobox_selecionar_tipo.get()
    numerodocumento = entry_numerodocumento.get()
    data_criacao = dt.datetime.now()
    data_criacao = data_criacao.strftime("%d/%m/%Y %H:%M")
    # Cadastrar em uma lista (necessário depois modificar para cadastrar em um banco de dados)
    clientes = len(lista_clientes)+1
    clientes_str = "Cliente Número-{}".format(clientes)
    lista_clientes.append((clientes_str,nomecliente,tipo,numerodocumento,data_criacao))

    comando = f"""DELETE FROM [dbo].[CADASTRO_CLIENTES]
    WHERE [Numero do Documento] = '{numerodocumento}'"""
    cursor.execute(comando)
    cursor.commit()


#Título da Janela

janela.title('Ferramenta de Cadastro de Novos Clientes')

label_nomecliente = tk.Label(text="Nome do Cliente")
label_nomecliente.grid(row=2, column=0,padx = 10, pady=10, sticky='nswe', columnspan =2 )

entry_nomecliente = tk.Entry()
entry_nomecliente.grid(row=2,column=2, padx=10, pady=10, sticky='nswe', columnspan = 4)

label_tipo_documento = tk.Label(text="Tipo do Documento do Cliente")
label_tipo_documento.grid(row=3, column=0,padx = 10, pady=10, sticky='nswe', columnspan =2 )

combobox_selecionar_tipo = ttk.Combobox(values=lista_documentos)
combobox_selecionar_tipo.grid(row=3, column=2, padx = 10, pady=10, sticky='nswe', columnspan = 2)

label_numerodocumento =  tk.Label(text="Digite Seu Documento")
label_numerodocumento.grid(row=4, column=0,padx = 10, pady=10, sticky='nswe', columnspan =2 )

entry_numerodocumento =  tk.Entry()
entry_numerodocumento.grid(row=4, column=2,padx = 10, pady=10, sticky='nswe', columnspan =2 )

botao_criar_codigo = tk.Button(text="Cadastrar Cliente", command=cadastrar_cliente)
botao_criar_codigo.grid(row=5,column=0,padx = 10, pady=10,sticky='nswe', columnspan =4)

botao_deletar_codigo = tk.Button(text="Deletar Cliente", command=deletar_cliente)
botao_deletar_codigo.grid(row=6,column=0,padx = 10, pady=10,sticky='nswe', columnspan =4)

janela.mainloop()

print(lista_clientes)