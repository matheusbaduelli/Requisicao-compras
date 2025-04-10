from purchasingreport.teste import test
from media_kanban.media import media_pecas,media_materia_prima

import pandas as pd
test()
requisicao = pd.read_excel("Estoque Almox/requisição.xlsx")
carteira = pd.read_excel("purchasingreport/resultado.xlsx")

requisicao1 = pd.DataFrame(requisicao)
carteira1 = pd.DataFrame(carteira)

carteira1["grupoConcatenado"] = carteira1["grupoConcatenado"].astype(str)
listaGP = []
listaCM = []
for linha in requisicao1["Gr"]:
    for linha2 in requisicao1["Codigo Material"]:
        listaGP.append(linha)
        listaCM.append(linha2)

media_materia_prima1 = pd.DataFrame(media_materia_prima(listaGP,listaCM))
media_pecas1 = pd.DataFrame(media_pecas(listaGP,listaCM))

# print(media_materia_prima1[["somaDasSomas","grupo_concatenado"]])
print(media_pecas1[["media_kanban1","grupo_concatenado"]])

requisicao1["Gr Concatenado"] = requisicao1["Gr"].astype(str) + requisicao1["Codigo Material"].astype(str)

requis = pd.merge(requisicao1,carteira1[["grupoConcatenado","somaComponente","somaMaterial"]],left_on="Gr Concatenado",right_on="grupoConcatenado",how="left")

requis.rename(columns={"somaComponente":"Qtd Pedido","somaMaterial":"Qtd Material carteira"},inplace=True)
requis.drop_duplicates(subset=["Gr Concatenado"],inplace=True)
requis.to_excel("Estoque Almox/result.xlsx",index=False)

# print(requis)