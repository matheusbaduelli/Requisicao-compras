import pandas as pd
import sqlite3
import os

from media_kanban.funcoes import lista_materia_prima,lista_pecas


def media_pecas(grupo,codigo):

  estrutura_pd1 = pd.read_excel('purchasingreport\Estrutura_Produtos_02_04_2025.xls')
  estrutura_pd2 = pd.read_excel('purchasingreport\Estrutura_Produtos_02_04_20252.xls')

  media_kanban_df = pd.read_excel('media_kanban\Média Kanban Key.xlsx')

  conn = sqlite3.connect('banco.db')
  cursor = conn.cursor()
  estrutura_concatenada = pd.concat([estrutura_pd1, estrutura_pd2])

  colunas_peso = ["Pes.Bruto", "Pes.Bruto.1", "Pes.Bruto.2", "Pes.Bruto.3","Qtde","Qtde.1","Qtde.2","Qtde.3"]

  for col in colunas_peso:
      estrutura_concatenada[col] = estrutura_concatenada[col].astype(str).str.replace(",", ".").astype(float)

  colunas_peso = ["Média arredondada"]

  for col in colunas_peso:
      media_kanban_df[col] = media_kanban_df[col].astype(str).str.replace(",", ".").astype(float)


  estrutura_concatenada.to_sql('estrutura_produtos', conn, if_exists='replace', index=False)

  media_kanban_df.to_sql('media_kanban', conn, if_exists='replace', index=False)

  query = """
  SELECT
    ep."Código do Produto",
    ep."Descrição do produto",
    ep."Grupo N1",
    ep."Código N1",
    ep."Qtde",
    ep."Descrição N1",
    ep."Grupo N2",
    ep."Código N2",
    ep."Qtde.1",
    ep."Descrição N2",
    ep."Grupo N3",
    ep."Código N3",
    ep."Qtde.2",
    ep."Descrição N3",
    ep."Grupo N4",
    ep."Código N4",
    ep."Qtde.3",
    ep."Descrição N4",
    ep."Pes.Bruto",
    ep."Pes.Bruto.1",
    ep."Pes.Bruto.2",
    ep."Pes.Bruto.3",
    mk."Média arredondada"
  FROM 
    estrutura_produtos ep
  INNER JOIN
    media_kanban mk
  ON
    ep."Código do Produto" = mk.Codigo
  """

  cursor.execute(query)
  dados = cursor.fetchall()



  de_resultado = pd.DataFrame(lista_pecas(dados,grupo,codigo))
  

  

  de_resultado["concatenado"] = de_resultado["Produto"].astype(str) + de_resultado["GrupoN1"].astype(str) + de_resultado["CodigoN1"].astype(str) + de_resultado["GrupoN2"].astype(str) + de_resultado["CodigoN2"].astype(str) + de_resultado["GrupoN3"].astype(str) + de_resultado["CodigoN3"].astype(str) + de_resultado["GrupoN4"].astype(str) + de_resultado["CodigoN4"].astype(str)
  de_resultado["grupo_concatenado"] = de_resultado["GrupoN1"].astype(str) + de_resultado["CodigoN1"].astype(str)
  
  de_resultado = de_resultado.drop_duplicates(subset='concatenado')

  # try:
  #     de_resultado.to_excel('media_kanban\dadospeca.xlsx', index=False)
      
  #     print("Arquivo Excel gerado com sucesso!")
  # except Exception as e:
  #     print(f"Erro ao executar consulta: {e}")
      
  # os.startfile("dados.xlsx")
  conn.close()
  return de_resultado

def media_materia_prima(grupo,codigo):

  estrutura_pd1 = pd.read_excel('purchasingreport\Estrutura_Produtos_02_04_2025.xls')
  estrutura_pd2 = pd.read_excel('purchasingreport\Estrutura_Produtos_02_04_20252.xls')

  media_kanban_df = pd.read_excel('media_kanban\Média Kanban Key.xlsx')

  conn = sqlite3.connect('banco.db')
  cursor = conn.cursor()
  estrutura_concatenada = pd.concat([estrutura_pd1, estrutura_pd2])

  colunas_peso = ["Pes.Bruto", "Pes.Bruto.1", "Pes.Bruto.2", "Pes.Bruto.3","Qtde","Qtde.1","Qtde.2","Qtde.3"]

  for col in colunas_peso:
      estrutura_concatenada[col] = estrutura_concatenada[col].astype(str).str.replace(",", ".").astype(float)

  colunas_peso = ["Média arredondada"]

  for col in colunas_peso:
      media_kanban_df[col] = media_kanban_df[col].astype(str).str.replace(",", ".").astype(float)


  estrutura_concatenada.to_sql('estrutura_produtos', conn, if_exists='replace', index=False)

  media_kanban_df.to_sql('media_kanban', conn, if_exists='replace', index=False)

  query = """
  SELECT
    ep."Código do Produto",
    ep."Descrição do produto",
    ep."Grupo N1",
    ep."Código N1",
    ep."Qtde",
    ep."Descrição N1",
    ep."Grupo N2",
    ep."Código N2",
    ep."Qtde.1",
    ep."Descrição N2",
    ep."Grupo N3",
    ep."Código N3",
    ep."Qtde.2",
    ep."Descrição N3",
    ep."Grupo N4",
    ep."Código N4",
    ep."Qtde.3",
    ep."Descrição N4",
    ep."Pes.Bruto",
    ep."Pes.Bruto.1",
    ep."Pes.Bruto.2",
    ep."Pes.Bruto.3",
    mk."Média arredondada"
  FROM 
    estrutura_produtos ep
  INNER JOIN
    media_kanban mk
  ON
    ep."Código do Produto" = mk.Codigo
  """

  cursor.execute(query)
  dados = cursor.fetchall()



  
  df_resultado = pd.DataFrame(lista_materia_prima(dados,grupo,codigo))

  df_resultado["somaDasSomas"] = df_resultado[["media_kanban1","media_kanban2","media_kanban3"]].sum(axis=1)
  df_resultado["grupo_concatenado"] = df_resultado["GrupoN1"].astype(str) + df_resultado["CodigoN1"].astype(str)

  df_resultado["concatenado"] = df_resultado["Produto"].astype(str) + df_resultado["GrupoN1"].astype(str) + df_resultado["CodigoN1"].astype(str) + df_resultado["GrupoN2"].astype(str) + df_resultado["CodigoN2"].astype(str) + df_resultado["GrupoN3"].astype(str) + df_resultado["CodigoN3"].astype(str) + df_resultado["GrupoN4"].astype(str) + df_resultado["CodigoN4"].astype(str)

  

  df_resultado = df_resultado.drop_duplicates(subset='concatenado')
  

  # try:
      
  #     df_resultado.to_excel('media_kanban\dados.xlsx', index=False)
  #     print("Arquivo Excel gerado com sucesso!")
  # except Exception as e:
  #     print(f"Erro ao executar consulta: {e}")
      
  # # os.startfile("dados.xlsx")
  conn.close()
  return df_resultado