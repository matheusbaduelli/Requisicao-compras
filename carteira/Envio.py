import pandas as pd
import math

def calcular_peso_bruto():

    # Ler a planilha
    planilha1 = pd.read_excel('purchasingreport\Estrutura_Produtos_02_04_2025.xls',usecols="A:AL")
    planilha2 = pd.read_excel('purchasingreport\Estrutura_Produtos_02_04_2025.xls',usecols="A:AL")

    # Concatenar as planilhas
    planilhaConcatenada = pd.concat([planilha1,planilha2],ignore_index=False)



    # Lista das colunas que devem ser convertidas para float
    colunas_peso = ["Pes.Bruto", "Pes.Bruto.1", "Pes.Bruto.2", "Pes.Bruto.3","Qtde","Qtde.1","Qtde.2","Qtde.3"]

    # Converter colunas para float, substituindo erros (ex: strings com vírgulas)
    for col in colunas_peso:
        planilhaConcatenada[col] = planilhaConcatenada[col].astype(str).str.replace(",", ".").astype(float)
    
   
        
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
        # # grupos = [6,2]
        # # codigos = [5306]
        # if GrupoN1 in grupos or GrupoN2 in grupos or GrupoN3 in grupos or GrupoN4 in grupos:
            # if CódigoN1 in codigos or CódigoN2 in codigos or CódigoN3 in codigos or CódigoN4 in codigos:

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
    # df.to_excel("purchasingreport/resultado12.xlsx", index=False)


    # # Carregar o arquivo Excel
    # arquivo = "purchasingreport/resultado12.xlsx"
    # de = pd.read_excel(arquivo, engine="openpyxl")  # engine é necessário para arquivos .xlsx

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
    # dg = pd.DataFrame(de)
    # # Remover duplicatas com base na coluna "CódigoConcatenado"
    df.drop_duplicates(subset=["GrupoECódigo"], inplace=True)
    
    # # Salvar novamente no Excel
    # dg.to_excel("purchasingreport/resultado123.xlsx", index=False, engine="openpyxl")

    return df


    # listagem = {
    #     "colunaA": df["CódigodoProduto"],
    #     "colunaB": df["DescriçãodoProduto"],
    #     "colunaC": df["GrupoN1"],
    #     "colunaD": df["CódigoN1"],
    #     "colunaE": df["Qtd"],
    #     "colunaF": df["DescricaoN1"]
    # }

    # listagem2 = {
    #     "colunaA": df["CódigodoProduto"],
    #     "colunaB": df["DescriçãodoProduto"],
    #     "colunaC": df["GrupoN2"],
    #     "colunaD": df["CódigoN2"],
    #     "colunaE": df["Qtd1"],
    #     "colunaF": df["DescricaoN2"]
    # }

    # listagem3 = {
    #     "colunaA": df["CódigodoProduto"],
    #     "colunaB": df["DescriçãodoProduto"],
    #     "colunaC": df["GrupoN3"],
    #     "colunaD": df["CódigoN3"],
    #     "colunaE": df["Qtd2"],
    #     "colunaF": df["DescricaoN3"]
    # }

    # listagem4 = {
    #     "colunaA": df["CódigodoProduto"],
    #     "colunaB": df["DescriçãodoProduto"],
    #     "colunaC": df["GrupoN4"],
    #     "colunaD": df["CódigoN4"],
    #     "colunaE": df["Qt3"],
    #     "colunaF": df["DescricaoN4"]
    # }


    # # Transforma os dois em DataFrames
    # df1 = pd.DataFrame(listagem)
    # df2 = pd.DataFrame(listagem2)
    # df3 = pd.DataFrame(listagem3)
    # df4 = pd.DataFrame(listagem4)

    # # Junta os dois DataFrames um embaixo do outro
    # df_final = pd.concat([df1, df2,df3,df4], ignore_index=True)

    # # Salva no Excel
    # df_final.to_excel("purchasingreport/listaConcatenada.xlsx", index=False, engine="openpyxl")

            
