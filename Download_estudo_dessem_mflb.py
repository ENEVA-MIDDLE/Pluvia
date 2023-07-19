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
from Funcoes_API_prospec import executa_dessem
from Funcoes_API_prospec import download_compilado
from Funcoes_API_prospec import GetStatusOfStudy
from Funcoes_API_prospec import arquiva_dados_rodada_dessem
from Funcoes_API_prospec import deleta_linhas_base_prospec_dessem
from Funcoes_API_prospec import envia_email_python
import pendulum
import logging
import datetime
  

idStudy=4782

token = autenticar_prospec()


##Download Prospec com time e verificando o status do estudo(concluido ou não)
# time.sleep(20*60)
# while True:
#     logging.info('Verificando se o estudo já foi finalizado:')
#     print('Verificando se o estudo já foi finalizado:')
#     status = GetStatusOfStudy(token, idStudy)
#     print(status)
#     if status == 'Finished':
#         #Download Propec
#         pathdownload=r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\07.Dessem\02.Rodadas\2021\09\25\04.Download Estudos'
#         nome_arquivo='ResultadoDessem_'+ str(idStudy) + '.zip'
#         nome_arquivo_semzip='ResultadoDessem_'+ str(idStudy) 
#         download_compilado(token, idStudy, pathdownload, nome_arquivo)
#         #Unzip pasta
#         shutil.unpack_archive(pathdownload +'\\' + nome_arquivo, pathdownload +'\\' + nome_arquivo_semzip,'zip')
#         break
#     else:
#         logging.info('Estudo ainda não concluído, iniciando espera para nova verificação')
#         print('Estudo ainda não concluído, iniciando espera para nova verificação')
#         time.sleep(4*60)

##Download Prospec
pathdownload=r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\07.Dessem\02.Rodadas\2021\09\25\04.Download Estudos'
nome_arquivo='ResultadoDessem_'+ str(idStudy) + '.zip'
nome_arquivo_semzip='ResultadoDessem_'+ str(idStudy) 
caminho_saida_rodada= pathdownload +'\\' + nome_arquivo_semzip
now=datetime.datetime.now().strftime('%d/%m/%Y')
download_compilado(token, idStudy, pathdownload, nome_arquivo)
#Unzip pasta
shutil.unpack_archive(pathdownload +'\\' + nome_arquivo, pathdownload +'\\' + nome_arquivo_semzip,'zip')


#Arquivando na base (substitui se já houver aquele id de estudo lá)
deleta_linhas_base_prospec_dessem(idStudy)
time.sleep(5)
arquiva_dados_rodada_dessem(idStudy, now, caminho_saida_rodada)

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

