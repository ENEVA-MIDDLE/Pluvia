import shutil
import time
from Funcoes_API_prospec import cria_pasta_rodada_dessem
from Funcoes_API_prospec import cria_estudo
from Funcoes_API_prospec import autenticar_prospec
from Funcoes_API_prospec import generateDessemStudyDecks
from Funcoes_API_prospec import sendFileToStudy
from pathlib import Path
from Funcoes_API_prospec import download_decks_iniciais
from Funcoes_API_prospec import verifica_arquivos_dessem
from Funcoes_API_Pluvia import cria_pasta_local_temporaria
from Funcoes_API_prospec import copia_arquivos_rodada_dessem
from Funcoes_API_prospec import edita_entdados_deleta_blocos
from Funcoes_API_prospec import le_arquivos_prev_carga_dessem
from Funcoes_API_prospec import busca_deck_dessem
from Funcoes_API_prospec import sendFileToDeck
import pendulum
import logging

data_rodada_dessem = pendulum.yesterday(tz='America/Sao_Paulo')    #pendulum.today(tz='America/Sao_Paulo')  
ano_rodada_dessem = data_rodada_dessem.year
mes_rodada_dessem = data_rodada_dessem. month
dia_rodada_dessem = data_rodada_dessem.day
resposta = cria_pasta_rodada_dessem(data_dessem = data_rodada_dessem)
nome_log = data_rodada_dessem.format('YYYY.MM.DD') + '_log_rodada_dessem.log'
caminho_rodada_dessem = resposta['caminho']
logging.basicConfig(filename=caminho_rodada_dessem + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')
logging.info('Iniciando a preparação da rodada do Dessem do dia' + data_rodada_dessem.format('DD/MM/YYYY'))



dias_dessem = 2
copia_arquivos_rodada_dessem(data_rodada_dessem, caminho_rodada_dessem)
resposta = verifica_arquivos_dessem(data_rodada_dessem)

if not(resposta['Arquivos Corretos']):
    print('Arquivos pendentes para criação dos estudos, verificar Log')
    logging.info('Arquivos pendentes na pasta de criação de estudos')
    quit()
token = autenticar_prospec()
logging.info('Autenticado no prospec com sucesso')
#função que cria um estudo e atribui um IdStudy a ele
nome_estudo_dessem = data_rodada_dessem.format('YYYY_MM_DD') + '_DESSEM'
logging.info('Iniciando criação do estudo:' + nome_estudo_dessem)
IdStudy = cria_estudo (nome_estudo_dessem)
logging.info('Estudo criado no prospec:' + str(IdStudy))
caminho_rodada_dessem = resposta['Caminho rodada dessem']
nome_deck_dessem = resposta['Nome do Deck Dessem']

caminho_deck_dessem = Path(caminho_rodada_dessem + '\\01.Decks\\' + nome_deck_dessem)
logging.info('Iniciando envio de arquivos para o estudo:' + nome_deck_dessem)
sendFileToStudy(IdStudy, caminho_deck_dessem, nome_deck_dessem)
logging.info('Solicitada a geração de decks do estudo.')
generateDessemStudyDecks(IdStudy, ano_rodada_dessem, mes_rodada_dessem, dia_rodada_dessem, dias_dessem, nome_deck_dessem)
logging.info('Decks para o estudo criados com sucesso:' + str(IdStudy))
time.sleep(60)
caminho_download = cria_pasta_local_temporaria()
logging.info('Iniciando o download dos decks iniciais do estudo:' + str(IdStudy) + ' - pasta temporária:' + caminho_download)
resposta = download_decks_iniciais(token, IdStudy, caminho_download)
caminho_pasta_decks_iniciais = resposta['caminho deck pasta temporária']
nome_pasta_decks_iniciais = resposta['Nome pasta deck']
caminho_pasta_decks_iniciais_rede = caminho_rodada_dessem + '\\02.Prevs\\' + nome_pasta_decks_iniciais
logging.info('Decks iniciais baixados no caminho:' + caminho_pasta_decks_iniciais)
shutil.move(caminho_pasta_decks_iniciais, caminho_rodada_dessem + '\\02.Prevs' )
time.sleep(15)
logging.info('Pasta descompactada com decks iniciais movida para a rede no caminho:' + caminho_pasta_decks_iniciais_rede)
shutil.rmtree(caminho_download)
logging.info('Pasta temporária deletada com sucesso no caminho' + caminho_download)

logging.info('Iniciando edição do arquivo entdados do primeiro dia do dessem')
caminho_entdados_editado = edita_entdados_deleta_blocos(data_rodada_dessem, caminho_pasta_decks_iniciais_rede)
logging.info('Arquivo entdados editado com sucesso (DP e DE) e salvo no caminho:' + caminho_entdados_editado)
print(caminho_entdados_editado)
logging.info('Iniciando leitura de arquivos DP.txt e DE.txt')
resposta = le_arquivos_prev_carga_dessem (caminho_rodada_dessem + '\\02.Prevs')
novo_bloco_DP = resposta['bloco_DP']
novo_bloco_DE = resposta['bloco_DE']
logging.info('Arquivos lidos com sucesso (DP.txt e DE.txt)')
logging.info('Abrindo arquivo entdados para inclusão dos blocos DP e DE da previsão de carga dessem')
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
logging.info('Arquivo entdados editado com sucesso e pronto para subida no deck do dessem no prospec')

#Inicia envio de arquivos para deck do dia seguinte do dessem

#pega o deck ID do segundo dia do dessem
resposta = busca_deck_dessem(IdStudy, data_rodada_dessem.add(days=1))
deckId_segundo_dia = resposta['Deck ID']

#envia os arquivos para o segundo dia do dessem
#arquivos a serem enviados:
#entdados.dat
#operut.dat
#dadvaz.dat
arquivos_upload_dessem = ['entdados.dat', 'operut.dat', 'dadvaz.dat']
for arquivo in arquivos_upload_dessem:
    caminho_arquivo = caminho_rodada_dessem + '\\02.Prevs\\' + arquivo
    sendFileToDeck(IdStudy, deckId_segundo_dia, caminho_arquivo, arquivo)
    print('Arquivo enviado para o estudo', IdStudy, 'Deck Id ', deckId_segundo_dia, ':', caminho_arquivo)
    logging.info('Arquivo enviado para o estudo' +  str(IdStudy) + 'Deck Id '+ str(deckId_segundo_dia) + ': ' + caminho_arquivo)

print('Iniciando a execução do estudo')

#Iniciar a execução do estudo

#Monitorar o estudo

#Download estudo

#Processa estudo e cataloga na base
