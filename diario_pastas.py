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
    while a:
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
        print ('Iniciando verificação de disponibilidade de PREVS (', pasta_relatorio, ') - data e hora: ', now.strftime('%d/%m/%Y %H:%M %S') )
        logging.info('Iniciando verificação de disponibilidade de PREVS (' +  pasta_relatorio + ')')
        resposta = verifica_disponibilidade_PREVS (data_mapas, Preliminar=x)
        lista_mapas = resposta['Lista de Mapas']
        print(resposta)
        logging.info(resposta)
        a = not (resposta['download_prevs'])
        if resposta['download_prevs']:
            break
        else:
            logging.info('Um ou mais arquivos de PREVS estão indisponíveis')
            time.sleep(5*60)
    logging.info('Prevs disponíveis, iniciando criação de pasta padrão na rede')
    #cria pasta padrão da rodada na rede
    caminho_rodada = cria_pastas_rodada(tipo_rodada='Diario')
    logging.info('Pastas criadas na rede, iniciando coleta de arquivos padrão da semana (deck NEWAVE, deck DECOMP e Arquivos GEVAZP).')
    copia_arquivos_padrão_rodada_diaria (caminho_rodada)

    logging.info('Arquivos de Newave, decomp e GEVAZP copiados para pasta padrão')

    #define o caminho da pasta das prevs definitivas ou preliminares
    caminho_rodada_prevs = os.path.join(caminho_rodada, '02.Prevs', pasta_relatorio)
    #cria pasta de prevs com nome definitivo ou prelimnar, dependendo da hora
    os.makedirs(os.path.join(caminho_rodada_prevs, 'prevs'), exist_ok=True)
    logging.info('Criadas pastas da rodada')
    resposta = baixa_prevs_configurada(lista_mapas)
    caminho_local = resposta
    print(resposta)
    for arquivo in os.listdir(caminho_local):  # lê os arquivos que estão na pasta de download
        #print(arquivo)
        if arquivo.upper().endswith('-PREVS.ZIP'):  # exibe somente os arquvos de ENA
            print(arquivo)
            caminho_arquivo = os.path.realpath(caminho_local) + '\\' + arquivo
            caminho_pasta = caminho_arquivo.replace('-PREVS.zip', '')
            with ZipFile(caminho_arquivo, 'r') as zipObj:  # descompacta o zip na mesma pasta do zip
                zipObj.extractall(caminho_pasta)
                print('arquivo descompactado:', arquivo)
                logging.info('arquivo descompactado:' + arquivo)
            completa_prevs(caminho_pasta)
            os.remove(os.path.join(caminho_local, arquivo))
            time.sleep(5)
            nome_pasta = caminho_pasta[caminho_pasta.rfind('\\') + 1:]
            print(nome_pasta)
            if nome_pasta.startswith('ONS'):
            #se o nome da pasta começa com ONS, renomeia os arquivos
                renomeia_prevs(caminho_pasta)
            print('caminho local: ',caminho_local)
            print('nome pasta:', nome_pasta)
            print('caminho rodada:', caminho_rodada_prevs)
            print('nome pasta:', nome_pasta)

            shutil.move(Path.joinpath(Path(caminho_local), nome_pasta), Path.joinpath(Path(caminho_rodada_prevs), nome_pasta))
            logging.info('pasta movida para a rede:' + nome_pasta)
    for folder in os.listdir(caminho_rodada_prevs):
        origem = os.path.join(caminho_rodada_prevs, folder)
        if folder != 'prevs':
            for file in os.listdir(origem):

                shutil.copy2(os.path.join(origem, file), os.path.join(caminho_rodada_prevs, 'prevs', file))

    limpa_prevs_semana_atual (caminho_rodada_prevs + r'\prevs')
    logging.info('Deletados arquivos de PREVS de sensibilidade da Semana Operativa vigente')

##NEW#pasta prevs individual por modelo
    for folder in os.listdir(caminho_rodada_prevs):
        print(os.listdir(caminho_rodada_prevs))
        origem = os.path.join(caminho_rodada_prevs, folder)
        prevsM1=r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\05.Arquivos Semanais\04.Prevs\prevsM1'#new
        if folder != 'prevs':
            for file in os.listdir(origem):
                subpasta=folder.partition("-")[0]
                os.makedirs(os.path.join(caminho_rodada_prevs, subpasta), exist_ok=True)
                shutil.copy2(os.path.join(origem, file), os.path.join(caminho_rodada_prevs, subpasta, file))
            #renomeia os prevs aaaamm-prevs.rvx
            caminho_renomeiaprevs = os.path.realpath(caminho_rodada_prevs) + '\\' + subpasta
            renomeia_prevs(caminho_renomeiaprevs)
            #Copiar prevs do M1 da pasta prevsM1
            for file in os.listdir(prevsM1): #new
                shutil.copy2(os.path.join(prevsM1, file), os.path.join(caminho_rodada_prevs, subpasta, file))#new
    
    logging.info('Download de PREVS concluído')

    caminho_arquivo_criacao_estudo = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\Criacao_Estudos_diario.xlsm'
    caminho = os.path.realpath(caminho_rodada)
    #gera_arquivo_UH_atualizado (caminho_rodada=caminho)#comentado porque tava dando erro no definitivo - arq aberto em outra maquina
    logging.info('Arquivo de volume UHE criado')
except:    
    logging.exception("ERRO DE CÓDIGO")
