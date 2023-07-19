#codigo temporario para ajudar a subir na base Prospec os estudos diarios já rodados e baixados

from Funcoes_API_prospec import cria_pastas_rodada
from Funcoes_API_Pluvia import baixa_prevs_configurada
from Funcoes_API_Pluvia import completa_prevs
from Funcoes_API_Pluvia import verifica_disponibilidade_PREVS
from Funcoes_API_prospec import renomeia_prevs
from Funcoes_API_prospec import limpa_prevs_semana_atual
from Funcoes_API_prospec import cria_rodada
from Funcoes_API_prospec import gera_arquivo_UH_atualizado
from Funcoes_API_prospec import copia_arquivos_padrão_rodada_diaria
from Funcoes_API_prospec import download_estudos_finalizados_diario
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

################ INPUT #####################
    StudyId=6613
    caminho_estudo_descompactado = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\2022\01\17\04.Download Estudos\Definitivo\1.0_2022117_REV2_JAN-22_A_JAN-22_DIARIO_DEFINITIVO_6613'
############################################    
    
    
    
    arquiva_dados_rodada_diaria(StudyId, data_mapas,caminho_estudo_descompactado, pasta_relatorio.upper())
    logging.info('Estudo arquivado na base de dados')
    
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