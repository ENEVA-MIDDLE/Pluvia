#Codigo 18/01/2022 17h45
#Coloca o estudo diario para rodar, monitora, baixa e joga os resultados para a base Prospec (fonte do Power BI)
#A parte do código que gerava relatório pdf a partir de um excel foi comentada e foi necessário adicionar algumas linhas de código
from Funcoes_API_prospec import cria_pastas_rodada
from Funcoes_API_Pluvia import baixa_prevs_configurada
from Funcoes_API_Pluvia import completa_prevs
from Funcoes_API_Pluvia import verifica_disponibilidade_PREVS
from Funcoes_API_prospec import renomeia_prevs
from Funcoes_API_prospec import limpa_prevs_semana_atual
from Funcoes_API_prospec import cria_rodada
from Funcoes_API_prospec import gera_arquivo_UH_atualizado
from Funcoes_API_prospec import copia_arquivos_padrão_rodada_diaria
from Funcoes_API_prospec import download_estudos_finalizados_diario2
from Funcoes_API_prospec import processa_rodada_diaria
from Funcoes_API_prospec import arquiva_dados_rodada_diaria
import datetime
import time
import os
from zipfile import ZipFile
import shutil
from pathlib import Path
import logging
from Funcoes_API_prospec import envia_email_python
nome_log = datetime.datetime.today().strftime('%Y.%m.%d') + '_log_prevs_diaria.log'
caminho_log = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\01.Log'
logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')
a = True
try:
    
    print('-----------------------------------------------------------------------')
    now = datetime.datetime.now()
    data_mapas = now.strftime('%d/%m/%Y')
    hora_limite = now.replace(hour=13, minute=00)

    if now < hora_limite:
        pasta_relatorio = 'Preliminar'
        x = True
    else:
        pasta_relatorio = 'Definitivo'
        x = False
 
    while True:
        logging.info('Iniciando verificação se relatório diário está concluído')
        resposta = download_estudos_finalizados_diario2 (pasta_relatorio)#baixa estudos, salva na rede, descompacta, joga na base prospec
        estudos_pendentes = resposta['Estudos pendentes']
        print(estudos_pendentes)
        if estudos_pendentes == 0:
            logging.info('Relatório diário concluído, inciando processamento')
            #caminho_zip = resposta['Caminho ZIP']
            #Chama as macros do excel para gerar relatório final em excel/pdf
            #resposta_procss = processa_rodada_diaria(caminho_zip, pasta_relatorio)
            logging.info('Relatório diário processado')
            print('Relatório diário processado')
            break
        else:
            logging.info('Relatório ainda não concluído, inciando espera para nova verificação')
            time.sleep(4*60)
    


#Processo add necessario para preencher base Prospec sem rodar as macros do relatório excel/pdf####
    #Caminho do estudo descompactado
    #caminho_estudo_descompactado = caminho_zip.replace('.zip','')
    #Descompactar o zipper baixado
    #shutil.unpack_archive(caminho_zip, caminho_estudo_descompactado,'zip')
    #Pegar id do estudo do nome da pasta descompactada
    #StudyId = int(caminho_estudo_descompactado[caminho_estudo_descompactado.rfind('_') + 1:])

#############################################################################################

    #Preenchimento da Base Prospec
    #arquiva_dados_rodada_diaria(StudyId, data_mapas,caminho_estudo_descompactado, pasta_relatorio.upper())
    #logging.info('Estudo arquivado na base de dados')

#############################################################################################    
    #envia arquivo da rodada diária por e-mail para substituir a base corrente no onedrive
    caminho_pasta_base_prospec = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Prospec'
    nome_base_prospec = 'Base_prospec_diario.xlsx'
    assunto = 'DIARIO - BASE PROSPEC PARA ATUALIZAR NO POWER BI'# + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
    corpo = 'Prezados,\n\nSegue em anexo base do prospec atualizada para update do Power Bi.\n\n'
    caminho_anexo = caminho_pasta_base_prospec + '\\' + nome_base_prospec
    anexo = nome_base_prospec
    destinatario = 'maria.barbosa@eneva.com.br' 
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
    logging.info('E-mail com base de dados do prospec enviado para ativação do fluxo do Power Automate')
except:
    logging.exception("ERRO DE CÓDIGO")
    assunto = 'ERRO DE CÓDIGO NO BOL. DIÁRIO - ' + pasta_relatorio + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
    corpo = 'Prezados,\n\nO relatório diário (' + pasta_relatorio + ') apresentou erro no processamento. Log do erro em anexo.\n\n'
    caminho_anexo = caminho_log + '\\' + nome_log
    anexo = nome_log
    destinatario = 'maria.barbosa@eneva.com.br'
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)