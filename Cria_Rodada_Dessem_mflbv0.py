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
from Funcoes_API_prospec import getIdOfServer
from Funcoes_API_prospec import runExecution
from Funcoes_API_prospec import download_compilado
from Funcoes_API_prospec import GetStatusOfStudy
from Funcoes_API_prospec import arquiva_dados_rodada_dessem
from Funcoes_API_prospec import deleta_linhas_base_prospec_dessem
from Funcoes_API_prospec import envia_email_python
from Funcoes_API_prospec import edita_operut
import pendulum
import logging
import datetime
import os
data_rodada_dessem = pendulum.today(tz='America/Sao_Paulo') # pendulum.yesterday(tz='America/Sao_Paulo')    #pendulum.today(tz='America/Sao_Paulo')  
ano_rodada_dessem = data_rodada_dessem.year
mes_rodada_dessem = data_rodada_dessem. month
dia_rodada_dessem = data_rodada_dessem.day
resposta = cria_pasta_rodada_dessem(data_dessem = data_rodada_dessem)
nome_log = data_rodada_dessem.format('YYYY.MM.DD') + '_log_rodada_dessem.log'
caminho_rodada_dessem = resposta['caminho']
logging.basicConfig(filename=caminho_rodada_dessem + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')
logging.info('Iniciando a preparação da rodada do Dessem do dia' + data_rodada_dessem.format('DD/MM/YYYY'))


##############################################################################################
#Cria pastas da rodada com os arquivos baixados (DE, DP, dadvaz, deck CCEE)
dias_dessem = 2
copia_arquivos_rodada_dessem(data_rodada_dessem, caminho_rodada_dessem)
resposta = verifica_arquivos_dessem(data_rodada_dessem)

if not(resposta['Arquivos Corretos']):
    print('Arquivos pendentes para criação dos estudos, verificar Log')
    logging.info('Arquivos pendentes na pasta de criação de estudos')
    quit()

##############################################################################################
#Cria rodada no Prospec
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
##############################################################################################
#Baixa decks de entrada gerados
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

##############################################################################################
#Edicao operut.dat e entdados.dat
logging.info('Iniciando edição do arquivo entdados do primeiro dia do dessem')
caminho_entdados_editado = edita_entdados_deleta_blocos(data_rodada_dessem.add(days=1), caminho_pasta_decks_iniciais_rede)
edita_operut(caminho_pasta_decks_iniciais_rede)
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
##############################################################################################
#Inicia envio de arquivos editados para os deck já criados no PROSPEC

#pega o deck ID do segundo dia do dessem
resposta = busca_deck_dessem(IdStudy, data_rodada_dessem.add(days=1))
deckId_segundo_dia = resposta['Deck ID']

#Envia os arquivos para o segundo dia do dessem
#arquivos a serem enviados:
#entdados.dat
#dadvaz.dat
arquivos_upload_dessem = ['entdados.dat', 'dadvaz.dat'] #['entdados.dat', 'operut.dat', 'dadvaz.dat']
for arquivo in arquivos_upload_dessem:
    caminho_arquivo = caminho_rodada_dessem + '\\02.Prevs\\' + arquivo
    sendFileToDeck(IdStudy, deckId_segundo_dia, caminho_arquivo, arquivo)
    print('Arquivo enviado para o estudo', IdStudy, 'Deck Id ', deckId_segundo_dia, ':', caminho_arquivo)
    logging.info('Arquivo enviado para o estudo' +  str(IdStudy) + 'Deck Id '+ str(deckId_segundo_dia) + ': ' + caminho_arquivo)

time.sleep(60)

#Envia os arquivos para o primeiro dia do dessem
#operut.dat
#pega o deck ID do primeiro dia do dessem
deckId_primeiro_dia = deckId_segundo_dia-1
arquivos_upload_dessem_inicial = ['operut.dat']
for arquivo in arquivos_upload_dessem_inicial:
   caminho_arquivo = caminho_rodada_dessem + '\\02.Prevs\\' + arquivo
   sendFileToDeck(IdStudy, deckId_primeiro_dia, caminho_arquivo, arquivo)
   print('Arquivo enviado para o estudo', IdStudy, 'Deck Id ', deckId_primeiro_dia, ':', caminho_arquivo)
   logging.info('Arquivo enviado para o estudo' +  str(IdStudy) + 'Deck Id '+ str(deckId_primeiro_dia) + ': ' + caminho_arquivo)

time.sleep(15)

##############################################################################################
#Executa o Dessem (basicamente dá o play)
print('Iniciando a execução do estudo')
idServer = getIdOfServer('Dessem')
runExecution(IdStudy, idServer,0, None,None,None,None,0,2,10)

##############################################################################################
#Monitora e faz o Download do estudo Prospec 
time.sleep(20*60)
while True:
    logging.info('Verificando se o estudo já foi finalizado:')
    print('Verificando se o estudo já foi finalizado:')
    status = GetStatusOfStudy(token, IdStudy)
    print(status)
    if status == 'Finished':
        ##Download Prospec
        pathdownload=os.path.join(caminho_rodada_dessem, '04.Download Estudos')
        nome_arquivo='ResultadoDessem_'+ str(IdStudy) + '.zip'
        nome_arquivo_semzip='ResultadoDessem_'+ str(IdStudy) 
        caminho_saida_rodada= pathdownload +'\\' + nome_arquivo_semzip
        now=datetime.datetime.now().strftime('%d/%m/%Y')
        download_compilado(token, IdStudy, pathdownload, nome_arquivo)
        #Unzip pasta
        shutil.unpack_archive(pathdownload +'\\' + nome_arquivo, pathdownload +'\\' + nome_arquivo_semzip,'zip')
        print('Arquivo unzipado')
        break
    else:
        logging.info('Estudo ainda não concluído, iniciando espera para nova verificação')
        print('Estudo ainda não concluído, iniciando espera para nova verificação')
        time.sleep(4*60)

##############################################################################################
#Download estudo 
#Arquivando na base (substitui se já houver aquele id de estudo lá)
print('Iniciando arquivamento na base')
deleta_linhas_base_prospec_dessem(IdStudy)
time.sleep(5)
#Processa estudo e cataloga na base
arquiva_dados_rodada_dessem(IdStudy, now, caminho_saida_rodada)
print('Informação arquivada')

##############################################################################################
#Envia por email para atualizar POWER BI
#envia arquivo da rodada dessem por e-mail para substituir a base corrente no onedrive (Fluxo Power Automate)
#a base do onedrive atualiza um POWER BI (Fluxo Power Automate)
caminho_pasta_base_prospecdessem = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Prospec'
nome_base_prospecdessem = 'Base_prospec_dessem.xlsx'
assunto = 'BASE PROSPEC DESSEM PARA ATUALIZAR NO POWER BI'# + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
corpo = 'Prezados,\n\nSegue em anexo base do prospec dessem atualizada para update do Power Bi.\n\n'
caminho_anexo = caminho_pasta_base_prospecdessem + '\\' + nome_base_prospecdessem
anexo = nome_base_prospecdessem
destinatario = 'maria.barbosa@eneva.com.br' 
envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
