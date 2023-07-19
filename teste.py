# import shutil
# import time
# from Funcoes_API_prospec import cria_pasta_rodada_dessem
# from Funcoes_API_prospec import cria_estudo
# from Funcoes_API_prospec import autenticar_prospec
# from Funcoes_API_prospec import generateDessemStudyDecks
# from Funcoes_API_prospec import sendFileToStudy
# from pathlib import Path
# from Funcoes_API_prospec import download_decks_iniciais
# from Funcoes_API_prospec import verifica_arquivos_dessem
# from Funcoes_API_Pluvia import cria_pasta_local_temporaria
# from Funcoes_API_prospec import copia_arquivos_rodada_dessem
# from Funcoes_API_prospec import edita_entdados_deleta_blocos
# from Funcoes_API_prospec import le_arquivos_prev_carga_dessem
# from Funcoes_API_prospec import busca_deck_dessem
# from Funcoes_API_prospec import sendFileToDeck
# from Funcoes_API_prospec import executa_dessem
# from Funcoes_API_prospec import download_compilado
# from Funcoes_API_prospec import GetStatusOfStudy
# from Funcoes_API_prospec import arquiva_dados_rodada_dessem
# from Funcoes_API_prospec import deleta_linhas_base_prospec_dessem
# from Funcoes_API_prospec import getListOfSpotInstancesTypes
# from Funcoes_API_prospec import executa_estudos
# from Funcoes_API_prospec import conta_requisicoes
# from Funcoes_API_prospec import getIdOfServer
# from Funcoes_API_prospec import runExecution
# import pendulum
# import logging
# import datetime


# token = autenticar_prospec()
# #getIdOfServer(Dessem)

# idStudy=int(4761)
# #idServer=int(1823)

# serverType='Dessem'

# ExecutionMode = 0  # Modo de execução(integer): 0 - Modo Pdrão, 1 - Consistência, 2 - Padrão + consistência
# InfeasibilityHandling = ''#2  # InfeasibilityHandling(integer): 0 - Parar estudo, 1 - Tratar inviabilidades, 2 - Ignorar inviabilidades, 3 - Tratar + Ignorar inviabilidades
# InfeasibilityHandlingSensibility = ''#2  # InfeasibilityHandlingSensibility(integer): 0 - Parar estudo, 1 - Tratar inviabilidades, 2 - Ignorar inviabilidades, 3 - Tratar + Ignorar inviabilidades
# maxRestarts = ''#10


# token = autenticar_prospec()
# qtd_reqs = conta_requisicoes(token)
# print('Quantidade de requisições utilizadas: ', qtd_reqs)

# #executa_dessem (idStudy, serverType, ExecutionMode, InfeasibilityHandling, InfeasibilityHandlingSensibility, maxRestarts)
# #executa_estudos(idStudy, serverType, ExecutionMode, InfeasibilityHandling, InfeasibilityHandlingSensibility, maxRestarts)
# idServer = getIdOfServer('Dessem')
# runExecution(idStudy, idServer,0, None,None,None,None,0,2,10)
##############################################################################################
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
import pendulum
import logging
import datetime
import os
data_rodada_dessem = pendulum.today(tz='America/Sao_Paulo') # pendulum.yesterday(tz='America/Sao_Paulo')    #pendulum.today(tz='America/Sao_Paulo')  
ano_rodada_dessem = data_rodada_dessem.year
mes_rodada_dessem = data_rodada_dessem. month
dia_rodada_dessem = data_rodada_dessem.day
resposta = verifica_arquivos_dessem(data_rodada_dessem)

token = autenticar_prospec()
IdStudy=4785

caminho_rodada_dessem = resposta['Caminho rodada dessem']
print(caminho_rodada_dessem)

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

#Download estudo
#Arquivando na base (substitui se já houver aquele id de estudo lá)
print('Iniciando arquivamento na base')
deleta_linhas_base_prospec_dessem(IdStudy)
time.sleep(5)
arquiva_dados_rodada_dessem(IdStudy, now, caminho_saida_rodada)
print('Informação arquivada')
#Processa estudo e cataloga na base


#Envia por email
#envia arquivo da rodada dessem por e-mail para substituir a base corrente no onedrive
caminho_pasta_base_prospecdessem = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Prospec'
nome_base_prospecdessem = 'Base_prospec_dessem.xlsx'
assunto = 'BASE PROSPEC DESSEM PARA ATUALIZAR NO POWER BI'# + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
corpo = 'Prezados,\n\nSegue em anexo base do prospec dessem atualizada para update do Power Bi.\n\n'
caminho_anexo = caminho_pasta_base_prospecdessem + '\\' + nome_base_prospecdessem
anexo = nome_base_prospecdessem
destinatario = 'maria.barbosa@eneva.com.br' 
envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
