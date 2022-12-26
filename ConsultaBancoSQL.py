import pyodbc

nome_site = input("Digite o nome do site:")

# Realizar conexão com banco de dados SQL
dados_conexao = (
    "Driver={SQL SERVER};"
    "Server=LucasMachine;"
    "Database=ALGAR_RELATORIO_GERAL;"
)

conexao = pyodbc.connect(dados_conexao)
print("Conexão bem sucedida")
cursor = conexao.cursor()

consulta = f"""SELECT * FROM  [dbo].['DADOS DETENTORAS$']
WHERE ESTACAO_SEM_CODIGO = '{nome_site}'"""

cursor.execute(consulta)

linhas = cursor.fetchall()

for linha in linhas:
    print("Nome do Site:", linha[0])
    print("Detentora do Site:", linha[1],"\n")




