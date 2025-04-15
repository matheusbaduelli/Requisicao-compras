import pandas as pd
import sqlite3
from Envio import calcular_peso_bruto

import os
# Chamar a função para calcular o peso bruto
calcular_peso_bruto()
# Carregar o arquivo Excel
planilha1 = pd.read_excel('resultado123.xlsx')


# Criar conexão com o banco de dados
lista_tecnica = planilha1
lista_pedidos = pd.read_excel('Lista de pedidos.xlsx')


# Criar conexão com o banco de dados
conn = sqlite3.connect('Lista.db')
cursor = conn.cursor()
# Criar tabelas
lista_tecnica.to_sql('lista_tecnica', conn, if_exists='replace', index=False)
lista_pedidos.to_sql('lista_de_pedidos',conn,if_exists='replace',index=False)



cursor.execute("SELECT lista_de_pedidos.Pedido,lista_tecnica.CódigodoProduto,lista_tecnica.DescriçãodoProduto,lista_tecnica.CódigoN1,lista_tecnica.DescricaoN1,lista_tecnica.Qtd,lista_tecnica.CódigoN2,lista_tecnica.DescricaoN2,lista_tecnica.Qtd1,lista_tecnica.CódigoN3,lista_tecnica.DescricaoN3,lista_tecnica.Qtd2,lista_tecnica.CódigoN4,lista_tecnica.DescricaoN4,lista_tecnica.Qt3,lista_tecnica.Somase,lista_tecnica.Somase1,lista_tecnica.Somase2,lista_de_pedidos.Qtd FROM lista_de_pedidos INNER JOIN lista_tecnica ON lista_de_pedidos.Codigo = lista_tecnica.CódigodoProduto;")



nova_lista = []
# Iterar sobre os resultados
for linha in cursor.fetchall():
    lista = {"Pedido":linha[0],"Produto":linha[1],"Descricao":linha[2],"CódigoN1":linha[3],"DescricaoN1":linha[4],"Qtd":linha[5],"CódigoN2":linha[6],"DescricaoN2":linha[7],"Qtd1":linha[8],"CódigoN3":linha[9],"DescricaoN3":linha[10],"Qtd2":linha[11],"CódigoN4":linha[12],"DescricaoN4":linha[13],"Qt3":linha[14],"QtdPedido":linha[18]*linha[5],"QtdMultiSomase": (linha[15] or 0) * (linha[18] or 0),
    "QtdMultiSomase1": (linha[16] or 0) * (linha[18] or 0),
    "QtdMultiSomase2": (linha[17] or 0) * (linha[18] or 0)
}
    nova_lista.append(lista)
# Criar DataFrame com os resultados
df =  pd.DataFrame(nova_lista)
# Exportar o DataFrame para um arquivo Excel
df.to_excel("resultado.xlsx",index=False)

os.startfile("resultado.xlsx")

conn.commit()
conn.close()