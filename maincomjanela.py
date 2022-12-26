import pandas as pd
import tkinter as tk
from tkinter import ttk

## Importando a planilha de lista de detentoras
planilha = pd.read_excel(r"C:\Users\Lucas\OneDrive\Área de Trabalho\Projetos Algar\Detentoras.xlsx")

lista_de_site = []
lista_de_site = planilha['SITE']
print(lista_de_site)

janela = tk.Tk()
janela.title("Analise Cadastro FSCIs")

label_selecione_site = tk.Label(text="Selecione o Nome do Site:")
label_selecione_site.grid(row=3, column=0, padx = 10, pady=10, sticky='nswe', columnspan =2 )

combobox_selecionar_tipo = ttk.Combobox(values=lista_de_site)
combobox_selecionar_tipo.grid(row=3, column=2, padx = 10, pady=10, sticky='nswe', columnspan = 2)

janela.mainloop()

'''def analise_cadastro(planilha, site_nova_FSCI):
    ## Criando uma lista para armazenar a coluna de sites cadastrados
    lista_de_site = []
    lista_de_site = planilha['SITE']

    incremento = 0
    for item in lista_de_site:
        if site_nova_FSCI in lista_de_site[incremento]:
            print("Site Cadastrado")
            global posicao_do_site_na_lista
            posicao_do_site_na_lista = incremento
            break
        incremento += 1
        if incremento == len(lista_de_site):
            print("Site não Cadastrado")

def consulta_dententora(planilha):
    print(planilha)
    lista_de_detentoras = []
    lista_de_detentoras = planilha['DETENTORA']
    nome_detentora = lista_de_detentoras[posicao_do_site_na_lista]
    print("A detentora do site é {}".format(nome_detentora))
    return nome_detentora'''






'''## Entrando com um dado a qual se deseja montar uma FSCI
site_nova_FSCI = str(input("Digite o site que deseja montar uma nova FSCI:"))
site_nova_FSCI = site_nova_FSCI.strip().upper()

analise_cadastro(planilha, site_nova_FSCI)
nome_detentora = consulta_dententora(planilha)

'''