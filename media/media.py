import pandas as pd
import sqlite3
import os

from media.funcoes import lista_materia_prima,lista_pecas


def media_pecas(grupo,codigo):

  estrutura_pd1 = pd.read_excel('carteira\Estrutura_Produtos_02_04_2025.xls')
  estrutura_pd2 = pd.read_excel('carteira\Estrutura_Produtos_02_04_20252.xls')

  media_kanban_df = pd.read_excel('media/Média Kanban Key.xlsx')

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


  lista1 = {
    "Produto":de_resultado["Produto"],
    "DescricaoProduto":de_resultado["DescricaoProduto"],
    "GrupoN1":de_resultado["GrupoN1"],
    "CodigoN1":de_resultado["CodigoN1"],
    "QtdN1":de_resultado["QtdN1"],
    "DescricaoN1":de_resultado["DescricaoN1"],
    "MediaArredondada":de_resultado["MediaArredondada"],
    "media_peca":de_resultado["media_peca"],
    
    }
  
  lista2 = {
    "Produto":de_resultado["Produto"],
    "DescricaoProduto":de_resultado["DescricaoProduto"],
    "GrupoN1":de_resultado["GrupoN2"],
    "CodigoN1":de_resultado["CodigoN2"],
    "QtdN1":de_resultado["QtdN2"],
    "DescricaoN1":de_resultado["DescricaoN2"],
    "MediaArredondada":de_resultado["MediaArredondada"],
    "media_peca":de_resultado["media_peca1"],
    }
  lista3 = {
    "Produto":de_resultado["Produto"],
    "DescricaoProduto":de_resultado["DescricaoProduto"],
    "GrupoN1":de_resultado["GrupoN3"],
    "CodigoN1":de_resultado["CodigoN3"],
    "QtdN1":de_resultado["QtdN3"],
    "DescricaoN1":de_resultado["DescricaoN3"],
    "MediaArredondada":de_resultado["MediaArredondada"],
    "media_peca":de_resultado["media_peca2"],
    }

  lista4 = {
    "Produto":de_resultado["Produto"],
    "DescricaoProduto":de_resultado["DescricaoProduto"],
    "GrupoN1":de_resultado["GrupoN4"],
    "CodigoN1":de_resultado["CodigoN4"],
    "QtdN1":de_resultado["QtdN4"],
    "DescricaoN1":de_resultado["DescricaoN4"],
    "MediaArredondada":de_resultado["MediaArredondada"],
    "media_peca":de_resultado["media_peca3"],
    }

  de_resultado = pd.concat([pd.DataFrame(lista1), pd.DataFrame(lista2), pd.DataFrame(lista3), pd.DataFrame(lista4)], ignore_index=True)
  

  # de_resultado["grupo_concatenado"] = de_resultado["GrupoN1"].astype(str) + de_resultado["CodigoN1"].astype(str)

  de_resultado["grupo_concatenado"] = (
  de_resultado["GrupoN1"].apply(lambda x: str(int(x)) if pd.notna(x) else "") +
  de_resultado["CodigoN1"].apply(lambda x: str(int(x)) if pd.notna(x) else "")
)
  de_resultado = de_resultado.drop_duplicates(subset=["Produto", "GrupoN1", "CodigoN1"])

  de_resultado = de_resultado[de_resultado["DescricaoN1"].notna()]

  de_resultado["media_peca"] = de_resultado.groupby([ "GrupoN1", "CodigoN1"])["media_peca"].transform("sum")


  conn.close()
  return de_resultado

def media_materia_prima(grupo,codigo):

  estrutura_pd1 = pd.read_excel('carteira/Estrutura_Produtos_02_04_2025.xls')
  estrutura_pd2 = pd.read_excel('carteira/Estrutura_Produtos_02_04_20252.xls')

  media_kanban_df = pd.read_excel('media/Média Kanban Key.xlsx')

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
  df_resultado.to_excel('media/dados2.xlsx', index=False)
  

  lista1 = {
    "Produto":df_resultado["Produto"],
    "DescricaoProduto":df_resultado["DescricaoProduto"],
    "GrupoN1":df_resultado["GrupoN2"],
    "CodigoN1":df_resultado["CodigoN2"],
    "QtdN1":df_resultado["QtdN2"],
    "DescricaoN1":df_resultado["DescricaoN2"],
    "MediaArredondada":df_resultado["MediaArredondada"],
    "media_kanban1":df_resultado["media_kanban1"],
    
    }
  lista2 = {
    "Produto":df_resultado["Produto"],
    "DescricaoProduto":df_resultado["DescricaoProduto"],
    "GrupoN1":df_resultado["GrupoN3"],
    "CodigoN1":df_resultado["CodigoN3"],
    "QtdN1":df_resultado["QtdN3"],
    "DescricaoN1":df_resultado["DescricaoN3"],
    "MediaArredondada":df_resultado["MediaArredondada"],
    "media_kanban1":df_resultado["media_kanban2"],
    }
  lista3 = {
    "Produto":df_resultado["Produto"],
    "DescricaoProduto":df_resultado["DescricaoProduto"],
    "GrupoN1":df_resultado["GrupoN4"],
    "CodigoN1":df_resultado["CodigoN4"],
    "QtdN1":df_resultado["QtdN4"],
    "DescricaoN1":df_resultado["DescricaoN4"],
    "MediaArredondada":df_resultado["MediaArredondada"],
    "media_kanban1":df_resultado["media_kanban3"],
    }
  
  df_resultado = pd.concat([pd.DataFrame(lista1), pd.DataFrame(lista2), pd.DataFrame(lista3)], ignore_index=True)
  df_resultado = df_resultado.drop_duplicates(subset=["Produto", "GrupoN1", "CodigoN1"])

  df_resultado = df_resultado[df_resultado["DescricaoN1"].notna()]
  df_resultado["media_materia_prima"] = df_resultado.groupby([ "GrupoN1", "CodigoN1"])["media_kanban1"].transform("sum")

  
  df_resultado["grupo_concatenado"] = (
  df_resultado["GrupoN1"].apply(lambda x: str(int(x)) if pd.notna(x) else "") +
  df_resultado["CodigoN1"].apply(lambda x: str(int(x)) if pd.notna(x) else "")
)


  conn.close()
  return df_resultado