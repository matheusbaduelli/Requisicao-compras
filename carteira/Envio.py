import pandas as pd


def calcular_peso_bruto():
    planilha1 = pd.read_excel('carteira\Estrutura_Produtos_02_04_2025.xls',usecols="A:AL")
    planilha2 = pd.read_excel('carteira\Estrutura_Produtos_02_04_20252.xls',usecols="A:AL")
    lista_pedidos = pd.read_excel('carteira\Lista de pedidos.xlsx')

    planilhaConcatenada = pd.concat([planilha1,planilha2],ignore_index=False)

    planilha = pd.DataFrame(planilhaConcatenada)

    colunas_peso = ["Pes.Bruto", "Pes.Bruto.1", "Pes.Bruto.2", "Pes.Bruto.3","Qtde","Qtde.1","Qtde.2","Qtde.3"]

    for col in colunas_peso:
        planilha[col] = planilha[col].astype(str).str.replace(",", ".").astype(float)
    
    lista1 = {
        "Código do Produto": planilha["Código do Produto"],
        "Descrição do produto": planilha["Descrição do produto"],
        "Descrição N1": planilha["Descrição N1"],
        "Grupo": planilha["Grupo N1"],
        "Código": planilha["Código N1"],
        "Qtde": planilha["Qtde"],
        "Pes.Bruto":planilha["Pes.Bruto"]
    }

    lista2 = {
        "Código do Produto": planilha["Código do Produto"],
        "Descrição do produto": planilha["Descrição do produto"],
        "Descrição N1": planilha["Descrição N2"],
        "Grupo": planilha["Grupo N2"],
        "Código": planilha["Código N2"],
        "Qtde": planilha["Qtde.1"],
        "Pes.Bruto":planilha["Pes.Bruto.1"]
    }
    lista3 = {
        "Código do Produto": planilha["Código do Produto"],
        "Descrição do produto": planilha["Descrição do produto"],
        "Descrição N1": planilha["Descrição N3"],
        "Grupo": planilha["Grupo N3"],
        "Código": planilha["Código N3"],
        "Qtde": planilha["Qtde.2"],
        "Pes.Bruto":planilha["Pes.Bruto.2"]
    }
    lista4 = {
        "Código do Produto": planilha["Código do Produto"],
        "Descrição do produto": planilha["Descrição do produto"],
        "Descrição N1": planilha["Descrição N4"],
        "Grupo": planilha["Grupo N4"],
        "Código": planilha["Código N4"],
        "Qtde": planilha["Qtde.3"],
        "Pes.Bruto":planilha["Pes.Bruto.3"]
    }

    
    listagem1 = pd.DataFrame(lista1)
    listagem2 = pd.DataFrame(lista2)
    listagem3 = pd.DataFrame(lista3)
    listagem4 = pd.DataFrame(lista4)

    resultado = pd.concat([listagem1, listagem2, listagem3, listagem4],ignore_index=True)

    planilha_pedidos = pd.merge(lista_pedidos[["Pedido", "Qtd","Codigo"]],resultado, left_on="Codigo", right_on="Código do Produto", how="left")

    planilha_pedidos["qtdXQtde"] = planilha_pedidos["Qtd"] * planilha_pedidos["Qtde"]
    planilha_pedidos["pesoXQtde"] = planilha_pedidos["Pes.Bruto"] * planilha_pedidos["Qtd"]

    planilha_pedidos.drop_duplicates(subset=["Código do Produto","Grupo","Código","Pedido"], inplace=True)

    planilha_pedidos["somaComponente"] = planilha_pedidos.groupby(["Grupo","Código"])["qtdXQtde"].transform("sum")
    planilha_pedidos["somaPeso"] = planilha_pedidos.groupby(["Grupo","Código"])["pesoXQtde"].transform("sum")

    planilha_pedidos = planilha_pedidos[planilha_pedidos["Descrição N1"].notna()]
    planilha_pedidos = planilha_pedidos[planilha_pedidos["Descrição N1"].astype(str).str.strip() != ""]


    planilha_pedidos["grupoConcatenado"] = (
    planilha_pedidos["Grupo"].apply(lambda x: str(int(x)) if pd.notna(x) else "") +
    planilha_pedidos["Código"].apply(lambda x: str(int(x)) if pd.notna(x) else "")
)



    planilha_pedidos.to_excel("carteira/resultado123.xlsx", index=False)

    

    return planilha_pedidos



# calcular_peso_bruto()
