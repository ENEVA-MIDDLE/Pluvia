from Funcoes_API_prospec import download_estudos_nao_finalizados_novo_arquivo
# excel com os estudos
caminho_rodada = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\2021\06\28'
nome_arquivo_estudos = 'Criacao_Estudos.xlsm'
caminho_excel = caminho_rodada + '\\' + nome_arquivo_estudos
#baixa estudo para pasta temporária
download_estudos_nao_finalizados_novo_arquivo(caminho_arquivo_estudos=caminho_excel)
print('arquivo baixado')
#verifica se já existe a pasta descompactada, se existir, deleta

#descompacta estudo

#arquiva estudo

#Deleta pasta descompactada


#deleta pasta temporária