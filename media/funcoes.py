def lista_pecas(dados,grupo,codigo):
    nova_lista = []

    for linha in dados:
        lista = {
            "Produto": linha[0],
            "DescricaoProduto": linha[1],
            "GrupoN1": linha[2],
            "CodigoN1": linha[3],
            "QtdN1": linha[4],
            "DescricaoN1": linha[5],
            "GrupoN2": linha[6],
            "CodigoN2": linha[7],
            "QtdN2": linha[8],
            "DescricaoN2": linha[9],
            "GrupoN3": linha[10],
            "CodigoN3": linha[11],
            "QtdN3": linha[12],
            "DescricaoN3": linha[13],
            "GrupoN4": linha[14],
            "CodigoN4": linha[15],
            "QtdN4": linha[16],
            "DescricaoN4": linha[17],
            "Pes.Bruto": linha[18],
            "Pes.Bruto1": linha[19],
            "Pes.Bruto2": linha[20],
            "Pes.Bruto3": linha[21],
            "MediaArredondada": linha[22],
            "media_peca": (linha[22] or 0) * (linha[4] or 0),
            "media_peca1": (linha[22] or 0) * (linha[8] or 0),
            "media_peca2": (linha[22] or 0) * (linha[12] or 0),
            "media_peca3": (linha[22] or 0) * (linha[16] or 0),
            
        }
        grupos = grupo
        codigos = codigo
        if linha[2] in grupos or linha[6] in grupos or linha[10] in grupos or linha[14] in grupos:
            if linha[3] in codigos or linha[7] in codigos or linha[11] in codigos or linha[15] in codigos:
                
        
                nova_lista.append(lista)
    return nova_lista


def lista_materia_prima(dados,grupo,codigo):
    nova_lista = []

    for linha in dados:
        lista = {
            "Produto": linha[0],
            "DescricaoProduto": linha[1],
            "GrupoN1": linha[2],
            "CodigoN1": linha[3],
            "QtdN1": linha[4],
            "DescricaoN1": linha[5],
            "GrupoN2": linha[6],
            "CodigoN2": linha[7],
            "QtdN2": linha[8],
            "DescricaoN2": linha[9],
            "GrupoN3": linha[10],
            "CodigoN3": linha[11],
            "QtdN3": linha[12],
            "DescricaoN3": linha[13],
            "GrupoN4": linha[14],
            "CodigoN4": linha[15],
            "QtdN4": linha[16],
            "DescricaoN4": linha[17],
            "Pes.Bruto": linha[18],
            "Pes.Bruto1": linha[19],
            "Pes.Bruto2": linha[20],
            "Pes.Bruto3": linha[21],
            "MediaArredondada": linha[22],
            "media_kanban1": (linha[22] or 0) * (linha[19] or 0),
            "media_kanban2": (linha[22] or 0) * (linha[20] or 0),
            "media_kanban3": (linha[22] or 0) * (linha[21] or 0),
            
        }
        grupos = grupo
        codigos = codigo
        if linha[2] in grupos or linha[6] in grupos or linha[10] in grupos or linha[14] in grupos:
            if linha[3] in codigos or linha[7] in codigos or linha[11] in codigos or linha[15] in codigos:
                
        
                nova_lista.append(lista)
    
    return nova_lista