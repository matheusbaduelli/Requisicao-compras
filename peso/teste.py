import pandas as pd
import sqlite3
from peso.Envio import calcular_peso_bruto

import os

def test(grupo,codigo):




    lista_tecnica = calcular_peso_bruto(grupo,codigo)
    lista_pedidos = pd.read_excel('peso\Lista de pedidos.xlsx')


   
    conn = sqlite3.connect('Lista.db')
    cursor = conn.cursor()
    
    lista_tecnica.to_sql('lista_tecnica', conn, if_exists='replace', index=False)
    lista_pedidos.to_sql('lista_de_pedidos',conn,if_exists='replace',index=False)



    cursor.execute("SELECT lista_de_pedidos.Pedido,lista_tecnica.CódigodoProduto,lista_tecnica.DescriçãodoProduto,lista_tecnica.CódigoN1,lista_tecnica.DescricaoN1,lista_tecnica.Qtd,lista_tecnica.CódigoN2,lista_tecnica.DescricaoN2,lista_tecnica.Qtd1,lista_tecnica.CódigoN3,lista_tecnica.DescricaoN3,lista_tecnica.Qtd2,lista_tecnica.CódigoN4,lista_tecnica.DescricaoN4,lista_tecnica.Qt3,lista_tecnica.Somase,lista_tecnica.Somase1,lista_tecnica.Somase2,lista_de_pedidos.Qtd,lista_tecnica.GrupoN1,lista_tecnica.GrupoN2,lista_tecnica.GrupoN3,lista_tecnica.GrupoN4 FROM lista_de_pedidos INNER JOIN lista_tecnica ON lista_de_pedidos.Codigo = lista_tecnica.CódigodoProduto;")



    nova_lista = []
    
    for linha in cursor.fetchall():
        lista = {"Pedido":linha[0],"Produto":linha[1],"Descricao":linha[2],"CódigoN1":linha[3],"DescricaoN1":linha[4],"Qtd":linha[5],"CódigoN2":linha[6],"DescricaoN2":linha[7],"Qtd1":linha[8],"CódigoN3":linha[9],"DescricaoN3":linha[10],"Qtd2":linha[11],"CódigoN4":linha[12],"DescricaoN4":linha[13],"Qt3":linha[14],"QtdPedido":linha[18]*linha[5],"QtdMultiSomase": (linha[15] or 0) * (linha[18] or 0),
        "QtdMultiSomase1": (linha[16] or 0) * (linha[18] or 0),
        "QtdMultiSomase2": (linha[17] or 0) * (linha[18] or 0),
        "GrupoN1":linha[19],
        "GrupoN2":linha[20],
        "GrupoN3":linha[21],
        "GrupoN4":linha[22],
    }
        nova_lista.append(lista)
    
    df =  pd.DataFrame(nova_lista)

    df["soma_se1"] = df.groupby(["GrupoN2","CódigoN2"])["QtdMultiSomase"].transform("sum")
    df["soma_se2"] = df.groupby(["GrupoN3","CódigoN3"])["QtdMultiSomase1"].transform("sum")
    df["soma_se3"] = df.groupby(["GrupoN4","CódigoN4"])["QtdMultiSomase2"].transform("sum")

    df["soma_se4"] = df[["soma_se1","soma_se2","soma_se3"]].sum(axis=1)

    lista1 = {"grupo":df["GrupoN2"],"codigo":df["CódigoN2"],"soma_se":df["soma_se1"]}
    lista2 = {"grupo":df["GrupoN3"],"codigo":df["CódigoN3"],"soma_se":df["soma_se2"]}
    lista3 = {"grupo":df["GrupoN4"],"codigo":df["CódigoN4"],"soma_se":df["soma_se3"]}

    listagem = pd.concat([pd.DataFrame(lista1),pd.DataFrame(lista2),pd.DataFrame(lista3)], ignore_index=True)

    listagem["codigo_concatenado"] = (listagem["grupo"].apply(lambda x: str(int(x)) if pd.notna(x) else "") +
    listagem["codigo"].apply(lambda x: str(int(x)) if pd.notna(x) else ""))

    listagem = listagem[listagem["codigo"].notna()]
    # listagem = listagem[listagem["codigo"].astype(str).str.strip() != ""]

    listagem.to_excel("peso/resultado.xlsx")
    conn.commit()
    conn.close()

    return listagem