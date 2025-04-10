import math
import pandas as pd
import sqlite3



# planilha1 = pd.read_excel('Estrutura_Produtos_30_01_2025.xls',usecols="A:AL")
# planilha2 = pd.read_excel('Estrutura_Produtos_30_01_20252.xls',usecols="A:AL")
planilha3 = pd.read_excel('resultado.xlsx',usecols="A:AL")


# planilhaConcatenada = pd.concat([planilha1,planilha2],ignore_index=False)



# planilhaConcatenada.to_excel('resultado.xlsx',index=False)


colunas = ["Código do Produto","Descrição do produto","Grupo N1",	"Prefixo N1"	,"Código N1"	,"Qtde",	"Pes.Bruto",	"Pes.Liq",	"Descrição N1",	"Grupo N2"	,"Prefixo N2","	Alma N2",	"Código N2",	"Qtde.1",	"Pes.Bruto.1",	"Pes.Liq.1",	"Und","	Descrição N2",	"Grupo N3",	"Prefixo N3",	"Alma N3",	"Código N3",	"Qtde.2",	"Pes.Bruto.2",	"Pes.Liq.2",	"Und.1",	"Descrição N3",	"Grupo N4",	"Prefixo N4",	"Alma N4",	"Código N4",	"Qtde.3",	"Pes.Bruto.3",	"Pes.Liq.3",	"Und.2",	"Descrição N4"]


# for i,col in enumerate(colunas):
#     print(f"posição: {i} coluna: {col}")

for CódigodoProduto,DescriçãodoProduto,CódigoN1,GrupoN1,GrupoN2,GrupoN3,GrupoN4,DescricaoN1,PesoBruto,PesBruto1 in zip(planilha3[colunas[0]],planilha3[colunas[1]],planilha3[colunas[4]],planilha3[colunas[2]],planilha3[colunas[9]],planilha3[colunas[18]],planilha3[colunas[27]],planilha3[colunas[8]],planilha3[colunas[14]],planilha3["Pes.Bruto.1"]):
    if isinstance(PesoBruto, float) and math.isnan(PesoBruto):
        PesoBruto = float(0)
    if type(PesoBruto) == str:
        valor_float = float(PesoBruto.replace(",", "."))
        lista = [CódigodoProduto,DescriçãodoProduto,GrupoN1,CódigoN1,GrupoN2,GrupoN3,GrupoN4,DescricaoN1,valor_float,float(PesBruto1)]
    else:
        lista = [CódigodoProduto,DescriçãodoProduto,GrupoN1,CódigoN1,GrupoN2,GrupoN3,GrupoN4,DescricaoN1,PesoBruto,float(PesBruto1)]
    
    grupos = [1,2,3,4,5,6,7,13]
    for listagem in grupos:
        if GrupoN1 == listagem or GrupoN2 == listagem or GrupoN3 == listagem or GrupoN4 == listagem:
            print([lista[0],lista[1],lista[2],lista[3],lista[4],lista[5],lista[6],lista[7],lista[8],lista[9]])
