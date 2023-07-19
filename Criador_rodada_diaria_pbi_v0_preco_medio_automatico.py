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
caminho_log = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\01.Log'
logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')

today = date.today()
# today = date(2022, 8, 15)

pld_semanal = get_pld_semanal(today=today)

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
    #if pasta_relatorio == 'Preliminar':
    copia_arquivos_padrão_rodada_diaria (caminho_rodada)
    print("a")  
    logging.info('Arquivos de Newave, decomp e GEVAZP copiados para pasta padrão')
    print("b")        
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
            #for file in os.listdir(prevsM1): #new
            #    shutil.copy2(os.path.join(prevsM1, file), os.path.join(caminho_rodada_prevs, subpasta, file))#new
    
    logging.info('Download de PREVS concluído')

    caminho_arquivo_criacao_estudo = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\Criacao_Estudos_diario.xlsm'
    caminho = os.path.realpath(caminho_rodada)
    #gera_arquivo_UH_atualizado (caminho_rodada=caminho)#comentado porque tava dando erro no definitivo - arq aberto em outra maquina
    logging.info('Arquivo de volume UHE criado')

  ######Cria rodada ######################################################################################################################## 
  
    cria_rodada (caminho_arquivo_criacao_estudo, Diario=pasta_relatorio, Caminho_rodada=caminho, Executa_rodada='True', Envia_volume_UHE='False' )
    logging.info('Rodada criada no prospec')
    #deck_decomp_semana_seguinte(idStudy)
    assunto = 'Estudo diário em execução no prospec - DIÁRIO ' + pasta_relatorio
    corpo = 'Prezados,\n\nUm estudo foi criado e está em execução no prospec com as configurações de PREVS (' + pasta_relatorio + ') contidas no caminho abaixo.\n\n' + caminho_rodada_prevs
    caminho_anexo = ''
    anexo = ''
    destinatario = 'maria.barbosa@eneva.com.br'
    #envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
    #logging.info('E-mail enviado confirmando a criação do estudo no prospec')
    #logging.info('Iniciando espera de 20 minutos para conclusão do estudo')
    time.sleep(15*60)
    while True:
        logging.info('Iniciando verificação se relatório diário está concluído')
        resposta = download_estudos_finalizados_diario2_com_mensal(pasta_relatorio, pld_semanal)#baixa estudos, salva na rede, descompacta, joga na base prospec
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
#    #envia arquivo da rodada diária por e-mail para substituir a base corrente no onedrive
#    caminho_pasta_base_prospec = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Prospec'
#    nome_base_prospec = 'Base_prospec_diario.xlsx'
#    assunto = 'DIARIO - BASE PROSPEC PARA ATUALIZAR NO POWER BI'# + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
#    corpo = 'Prezados,\n\nSegue em anexo base do prospec atualizada para update do Power Bi.\n\n'
#    caminho_anexo = caminho_pasta_base_prospec + '\\' + nome_base_prospec
#    anexo = nome_base_prospec
#    destinatario = 'maria.barbosa@eneva.com.br' 
#    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
#    logging.info('E-mail com base de dados do prospec enviado para ativação do fluxo do Power Automate')
except:
    logging.exception("ERRO DE CÓDIGO")
    assunto = 'ERRO DE CÓDIGO NO BOL. DIÁRIO - ' + pasta_relatorio + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
    corpo = 'Prezados,\n\nO relatório diário (' + pasta_relatorio + ') apresentou erro no processamento. Log do erro em anexo.\n\n'
    caminho_anexo = caminho_log + '\\' + nome_log
    anexo = nome_log
    destinatario = 'maria.barbosa@eneva.com.br'
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)