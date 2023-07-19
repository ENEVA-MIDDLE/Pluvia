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
from Funcoes_API_prospec import download_estudos_finalizados_diario2, download_estudos_finalizados_diario2_com_mensal
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
from datetime import date
from get_pld_semanal import get_pld_semanal


nome_log = datetime.datetime.today().strftime('%Y.%m.%d') + '_log_prevs_diaria.log'
caminho_log = r'C:\SCRIPTS_\Log'
    #r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\01.Log'
logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')

today = date.today()
# today = date(2022, 8, 15)
caminho_rodada = r'C:\Users\alex.lourenco\OneDrive - Eneva S.A\Documentos\processos_alex\diario'

pld_semanal = get_pld_semanal(today=today)

print('-----------------------------------------------------------------------')
now = datetime.datetime.now()
data_mapas = now.strftime('%d/%m/%Y')
hora_limite = now.replace(hour=16, minute=00)

if now < hora_limite:
    pasta_relatorio = 'Preliminar'
    x = True
else:
    pasta_relatorio = 'Definitivo'
    x = False

caminho_rodada_prevs = os.path.join(caminho_rodada, '02.Prevs', pasta_relatorio)

######Cria rodada ########################################################################################################################

# cria_rodada (caminho_arquivo_criacao_estudo, Diario=pasta_relatorio, Caminho_rodada=caminho, Executa_rodada='True', Envia_volume_UHE='False' )
logging.info('Rodada criada no prospec')
assunto = 'Estudo diário em execução no prospec - DIÁRIO ' + pasta_relatorio
corpo = 'Prezados,\n\nUm estudo foi criado e está em execução no prospec com as configurações de PREVS (' + pasta_relatorio + ') contidas no caminho abaixo.\n\n' + caminho_rodada_prevs
caminho_anexo = ''
anexo = ''
destinatario = 'alex.lourenco@eneva.com.br'
# time.sleep(15*60)
while True:
    logging.info('Iniciando verificação se relatório diário está concluído')
    resposta = download_estudos_finalizados_diario2_com_mensal(pasta_relatorio, pld_semanal)  # baixa estudos, salva na rede, descompacta, joga na base prospec
    estudos_pendentes = resposta['Estudos pendentes']
    print(estudos_pendentes)
    if estudos_pendentes == 0:
        logging.info('Relatório diário concluído, inciando processamento')
        # caminho_zip = resposta['Caminho ZIP']
        # Chama as macros do excel para gerar relatório final em excel/pdf
        # resposta_procss = processa_rodada_diaria(caminho_zip, pasta_relatorio)
        logging.info('Relatório diário processado')
        print('Relatório diário processado')
        break
    else:
        logging.info('Relatório ainda não concluído, inciando espera para nova verificação')
        time.sleep(4 * 60)