import pendulum
from Funcoes_API_prospec import edita_entdados_deleta_blocos
from Funcoes_API_prospec import le_arquivos_prev_carga_dessem

#DADOS FORNECIDOS DE OUTRAS ETAPAS DO PROCESSO
data_dessem = pendulum.today()

caminho_deck = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\07.Dessem\02.Rodadas\2021\08\04\02.Prevs\4520_decks_entrada'
caminho_rodada_dessem = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\07.Dessem\02.Rodadas\2021\08\04'
#----------------------------------------------------------------------------------------------

caminho_entdados_editado = edita_entdados_deleta_blocos(data_dessem, caminho_deck)

print(caminho_entdados_editado)

resposta = le_arquivos_prev_carga_dessem (caminho_rodada_dessem + '\\02.Prevs')
novo_bloco_DP = resposta['bloco_DP']
novo_bloco_DE = resposta['bloco_DE']

#Abre o arquivo entdados (já com os blocos DP e DE-parcial deletados e salva os dados em uma variável)
with open(caminho_entdados_editado, 'r') as arquivo_entdados:
    lista_entdados = arquivo_entdados.readlines()
    #print(lista_entdados)

#abre novamente o arquivo entdados editado e insere os blocos da previsão de carga do dessem (DP e DE)
with open(caminho_entdados_editado, 'w') as arquivo_entdados:
    entdados_DP_DE = []
    #loop que lê o arquivo e identifica as linhas marcadas por outra função como início dos blocos DP e DE
    for line in lista_entdados:
        if line == '&DP - INICIO BLOCO DE CARGA\n':
            entdados_DP_DE.extend(novo_bloco_DP)
        elif line == '&DE - INICIO BLOCO DE DEMANDAS/CARGAS ESPECIAIS\n':
            entdados_DP_DE.extend(novo_bloco_DE)
        else:
            entdados_DP_DE.append(line)
    arquivo_entdados.writelines(entdados_DP_DE)
