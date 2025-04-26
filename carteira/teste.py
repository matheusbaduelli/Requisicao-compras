import pandas as pd
import sqlite3
import os
from carteira.Envio import calcular_peso_bruto
# Chamar a função para calcular o peso bruto

def test():
    # calcular_peso_bruto()
    # # Carregar o arquivo Excel
    # planilha1 = pd.read_excel('purchasingreport/resultado123.xlsx')


    # Criar conexão com o banco de dados
    lista_tecnica = calcular_peso_bruto()
    lista_pedidos = pd.read_excel('carteira\Lista de pedidos.xlsx')


    # Criar conexão com o banco de dados
    conn = sqlite3.connect('Lista.db')
    cursor = conn.cursor()
    # Criar tabelas
    lista_tecnica.to_sql('lista_tecnica', conn, if_exists='replace', index=False)
    lista_pedidos.to_sql('lista_de_pedidos',conn,if_exists='replace',index=False)



    cursor.execute("SELECT lista_de_pedidos.Pedido,lista_tecnica.CódigodoProduto,lista_tecnica.DescriçãodoProduto,lista_tecnica.CódigoN1,lista_tecnica.DescricaoN1,lista_tecnica.Qtd,lista_tecnica.CódigoN2,lista_tecnica.DescricaoN2,lista_tecnica.Qtd1,lista_tecnica.CódigoN3,lista_tecnica.DescricaoN3,lista_tecnica.Qtd2,lista_tecnica.CódigoN4,lista_tecnica.DescricaoN4,lista_tecnica.Qt3,lista_tecnica.Somase,lista_tecnica.Somase1,lista_tecnica.Somase2,lista_de_pedidos.Qtd,lista_tecnica.GrupoN1,lista_tecnica.GrupoN2,lista_tecnica.GrupoN3,lista_tecnica.GrupoN4 FROM lista_de_pedidos INNER JOIN lista_tecnica ON lista_de_pedidos.Codigo = lista_tecnica.CódigodoProduto;")



    nova_lista = []
    # Iterar sobre os resultados
    for linha in cursor.fetchall():
        lista = {"Pedido":linha[0],"Produto":linha[1],"Descricao":linha[2],"CódigoN1":linha[3],"DescricaoN1":linha[4],"Qtd":linha[5],"CódigoN2":linha[6],"DescricaoN2":linha[7],"Qtd1":linha[8],"CódigoN3":linha[9],"DescricaoN3":linha[10],"Qtd2":linha[11],"CódigoN4":linha[12],"DescricaoN4":linha[13],"Qt3":linha[14],"QtdPedido":linha[18],"QtdMultiSomase": (linha[15] or 0) * (linha[18] or 0),
        "QtdMultiSomase1": (linha[16] or 0) * (linha[18] or 0),
        "QtdMultiSomase2": (linha[17] or 0) * (linha[18] or 0),
        "GrupoN1":linha[19],
        "GrupoN2":linha[20],
        "GrupoN3":linha[21],
        "GrupoN4":linha[22],
    }
        nova_lista.append(lista)
    # Criar DataFrame com os resultados
    df =  pd.DataFrame(nova_lista)

    nivel1 = {"grupon1":df["GrupoN1"].astype(str) + df["CódigoN1"].astype(str),"qtd":df["Qtd"],"qtd_pedido":df["QtdPedido"],"qtd_pedidoXQtd":df["QtdPedido"] * df["Qtd"]}
    nivel2 = {"grupon1":df["GrupoN2"].astype(str) + df["CódigoN2"].astype(str),"qtd":df["Qtd1"],"qtd_pedido":df["QtdPedido"],"qtd_pedidoXQtd":df["QtdPedido"] * df["Qtd1"]}
    nivel3 = {"grupon1":df["GrupoN3"].astype(str) + df["CódigoN3"].astype(str),"qtd":df["Qtd2"],"qtd_pedido":df["QtdPedido"],"qtd_pedidoXQtd":df["QtdPedido"] * df["Qtd2"]}
    nivel4 = {"grupon1":df["GrupoN4"].astype(str) + df["CódigoN4"].astype(str),"qtd":df["Qt3"],"qtd_pedido":df["QtdPedido"],"qtd_pedidoXQtd":df["QtdPedido"] * df["Qt3"]}
    
   
    
    

    n1 = pd.DataFrame(nivel1)
    n2 = pd.DataFrame(nivel2)
    n3 = pd.DataFrame(nivel3)
    n4 = pd.DataFrame(nivel4)
    

    df["somaDasSomas"] = df[["QtdMultiSomase","QtdMultiSomase1","QtdMultiSomase2"]].sum(axis=1)
    df["somaMaterial"] = df.groupby(["CódigoN1","GrupoN1"])["somaDasSomas"].transform("sum")
    df["somaComponente"] = df.groupby(["CódigoN1","GrupoN1"])["QtdPedido"].transform("sum")

    concatenar = pd.concat([n1,n2,n3,n4],ignore_index=True)
    
    concatenar["somaMaterial"] = concatenar.groupby(["grupon1"])["qtd_pedidoXQtd"].transform("sum")
    
    
    
    df["grupoConcatenado"] = df["GrupoN1"].astype(str) + df["CódigoN1"].astype(str)
    
    


    # Exportar o DataFrame para um arquivo Excel
    df.to_excel("carteira/resultado.xlsx",index=False)
    concatenar.to_excel("carteira/concatenar.xlsx",index=False)


    
    conn.commit()
    conn.close()

