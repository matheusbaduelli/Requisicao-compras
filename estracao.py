from carteira.teste import test
from media.media import media_pecas,media_materia_prima

import pandas as pd
test()
requisicao = pd.read_excel("Estoque Almox/requisição.xlsx")
carteira = pd.read_excel("carteira/resultado.xlsx")

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


media_material = pd.merge(media_materia_prima1[["grupo_concatenado","mediaMateriaPrima"]],media_pecas1[["media_peca","grupo_concatenado"]],left_on="grupo_concatenado",right_on="grupo_concatenado",how="left")



requisicao1["Gr Concatenado"] = requisicao1["Gr"].astype(str) + requisicao1["Codigo Material"].astype(str)

requis = pd.merge(requisicao1,carteira1[["grupoConcatenado","somaComponente","somaMaterial"]],left_on="Gr Concatenado",right_on="grupoConcatenado",how="left")

requis_material = pd.merge(requis,media_material,left_on="Gr Concatenado",right_on="grupo_concatenado",how="left")

requis.rename(columns={"somaComponente":"Qtd Pedido","somaMaterial":"Qtd Material carteira"},inplace=True)
requis_material.drop_duplicates(subset=["Gr Concatenado"],inplace=True)
requis_material.to_excel("Estoque Almox/result.xlsx",index=False)

# print(requis)