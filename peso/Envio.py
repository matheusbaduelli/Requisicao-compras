import pandas as pd
import math

def calcular_peso_bruto(grupo,codigo):

    # Ler a planilha
    planilha1 = pd.read_excel('peso\Estrutura_Produtos_02_04_2025.xls',usecols="A:AL")
    planilha2 = pd.read_excel('peso\Estrutura_Produtos_02_04_20252.xls',usecols="A:AL")

    # Concatenar as planilhas
    planilhaConcatenada = pd.concat([planilha1,planilha2],ignore_index=False)



    # Lista das colunas que devem ser convertidas para float
    colunas_peso = ["Pes.Bruto", "Pes.Bruto.1", "Pes.Bruto.2", "Pes.Bruto.3","Qtde","Qtde.1","Qtde.2","Qtde.3"]

    # Converter colunas para float, substituindo erros (ex: strings com vírgulas)
    for col in colunas_peso:
        planilhaConcatenada[col] = planilhaConcatenada[col].astype(str).str.replace(",", ".").astype(float)
    
    grupos = grupo
    codigos = codigo
        
    lista2 = []
    # Loop sobre as colunas
    for _, row in planilhaConcatenada.iterrows():
        CódigodoProduto = row["Código do Produto"]
        DescriçãodoProduto = row["Descrição do produto"]
        GrupoN1 = row["Grupo N1"]
        CódigoN1 = row["Código N1"]
        DescricaoN1 = row["Descrição N1"]
        Qtd = row["Qtde"]
        PesoBruto = row["Pes.Bruto"]
        GrupoN2 = row["Grupo N2"]
        CódigoN2 = row["Código N2"]
        DescricaoN2 = row["Descrição N2"]
        Qtd1 = row["Qtde.1"]
        PesoBruto1 = row["Pes.Bruto.1"]
        GrupoN3 = row["Grupo N3"]
        CódigoN3 = row["Código N3"]
        DescricaoN3 = row["Descrição N3"]
        Qtd2 = row["Qtde.2"]
        PesoBruto2 = row["Pes.Bruto.2"]
        GrupoN4 = row["Grupo N4"]
        CódigoN4 = row["Código N4"]
        DescricaoN4 = row["Descrição N4"]
        Qt3 = row["Qtde.3"]
        PesoBruto3 = row["Pes.Bruto.3"]
        MultiPesoBruto = PesoBruto1 * Qtd
        MultiPesoBruto1 = PesoBruto2 * Qtd1
        MultiPesoBruto2 = PesoBruto3 * Qtd2
    

        # Tratar valores NaN
        if isinstance(PesoBruto, float) and math.isnan(PesoBruto):
            PesoBruto = 0.0
        if isinstance(PesoBruto1, float) and math.isnan(PesoBruto1):
            PesoBruto1 = 0.0
        if isinstance(PesoBruto2, float) and math.isnan(PesoBruto2):
            PesoBruto2 = 0.0
        if isinstance(PesoBruto3, float) and math.isnan(PesoBruto2):
            PesoBruto3 = 0.0
        
        # Adicionar os valores à lista
        lista = {"CódigodoProduto":CódigodoProduto, "DescriçãodoProduto":DescriçãodoProduto, "GrupoN1":GrupoN1, "CódigoN1":CódigoN1, "DescricaoN1":DescricaoN1, "Qtd":Qtd, "PesoBruto":PesoBruto, "GrupoN2":GrupoN2,"CódigoN2":CódigoN2, "DescricaoN2":DescricaoN2, "Qtd1":Qtd1, "PesoBruto1":PesoBruto1, "GrupoN3":GrupoN3, "CódigoN3":CódigoN3, "DescricaoN3":DescricaoN3, "Qtd2":Qtd2, "PesoBruto2":PesoBruto2, "GrupoN4":GrupoN4, "CódigoN4":CódigoN4, "DescricaoN4":DescricaoN4, "Qt3":Qt3, "PesoBruto3":PesoBruto3,"MultiPesoBruto":MultiPesoBruto,"MultiPesoBruto1":MultiPesoBruto1,"MultiPesoBruto2":MultiPesoBruto2}

        
        
        
        # Verificar os grupos
        # grupos = [1, 2, 3, 4, 5, 6, 7, 13]

        if GrupoN1 in grupos or GrupoN2 in grupos or GrupoN3 in grupos or GrupoN4 in grupos:
            if CódigoN1 in codigos or CódigoN2 in codigos or CódigoN3 in codigos or CódigoN4 in codigos:

                lista2.append(lista)

    
    

    # Create a list of column names
    column_names = [
        "CódigodoProduto", "DescriçãodoProduto", "GrupoN1", "CódigoN1", 
        "DescricaoN1", "Qtd", "PesoBruto", "GrupoN2", "CódigoN2", 
        "DescricaoN2", "Qtd1", "PesoBruto1", "GrupoN3", "CódigoN3", 
        "DescricaoN3", "Qtd2", "PesoBruto2", "GrupoN4", "CódigoN4", 
        "DescricaoN4", "Qt3", "PesoBruto3", "MultiPesoBruto", 
        "MultiPesoBruto1", "MultiPesoBruto2"
    ]

    # Create DataFrame with named columns
    df = pd.DataFrame(lista2, columns=column_names)

    # Export to Excel with headers
    


    # Carregar o arquivo Excel
    # f

    # Criar uma nova coluna e adicionar valores
    df["SomaSe"] = df.groupby(["CódigodoProduto","GrupoN2","CódigoN2"])["MultiPesoBruto"].transform("sum")
    df["SomaSe1"] = df.groupby(["CódigodoProduto","GrupoN3","CódigoN3"])["MultiPesoBruto1"].transform("sum")
    df["SomaSe2"] = df.groupby(["CódigodoProduto","GrupoN4","CódigoN4"])["MultiPesoBruto2"].transform("sum")
    df["GrupoECódigo"] = (
    df["GrupoN1"].astype(str) + "-" +
    df["CódigoN1"].astype(str) + "-" +
    df["GrupoN2"].astype(str) + "-" +
    df["CódigoN2"].astype(str) + "-" +
    df["GrupoN3"].astype(str) + "-" +
    df["CódigoN3"].astype(str) + "-" +
    df["GrupoN4"].astype(str) + "-" +
    df["CódigoN4"].astype(str) + "-" +
    df["CódigodoProduto"].astype(str)
)
    


 
    
    # Criar um novo DataFrame com as colunas desejadas
    
    # Remover duplicatas com base na coluna "CódigoConcatenado"
    df.drop_duplicates(subset=["GrupoECódigo"], inplace=True)
    
    # Salvar novamente no Excel
    return df

    

            
# calcular_peso_bruto()