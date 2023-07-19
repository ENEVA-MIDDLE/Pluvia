from Funcoes_API_prospec import cria_pastas_rodada
from Funcoes_API_Pluvia import baixa_prevs_configurada
from Funcoes_API_Pluvia import completa_prevs
from Funcoes_API_Pluvia import verifica_disponibilidade_PREVS
from Funcoes_API_prospec import renomeia_prevs
from Funcoes_API_prospec import limpa_prevs_semana_atual
from Funcoes_API_prospec import cria_rodada
#from Funcoes_API_prospec import gera_arquivo_UH_atualizado
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
caminho_log = r'C:\Users\alex.lourenco\OneDrive - Eneva S.A\Documentos\processos_alex\diario\log'
    #r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\01.Log'
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
    logging.info('Download de PREVS concluído')
    caminho_arquivo_criacao_estudo = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\Criacao_Estudos_diario.xlsm'
    caminho = os.path.realpath(caminho_rodada)
    #gera_arquivo_UH_atualizado (caminho_rodada=caminho)#comentado porque tava dando erro no definitivo - arq aberto em outra maquina
    logging.info('Arquivo de volume UHE criado')
    cria_rodada (caminho_arquivo_criacao_estudo, Diario=pasta_relatorio, Caminho_rodada=caminho, Executa_rodada='True', Envia_volume_UHE='False' )
    logging.info('Rodada criada no prospec')
    #deck_decomp_semana_seguinte(idStudy)
    assunto = 'Estudo diário em execução no prospec - DIÁRIO ' + pasta_relatorio
    corpo = 'Prezados,\n\nUm estudo foi criado e está em execução no prospec com as configurações de PREVS (' + pasta_relatorio + ') contidas no caminho abaixo.\n\n' + caminho_rodada_prevs
    caminho_anexo = ''
    anexo = ''
    destinatario = 'renata.hunder@eneva.com.br; middle@eneva.com.br'
    #envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
    #logging.info('E-mail enviado confirmando a criação do estudo no prospec')
    #logging.info('Iniciando espera de 20 minutos para conclusão do estudo')
    time.sleep(20*60)
    while True:
        logging.info('Iniciando verificação se relatório diário está concluído')
        resposta = download_estudos_finalizados_diario (pasta_relatorio)
        estudos_pendentes = resposta['Estudos pendentes']
        if estudos_pendentes == 0:
            logging.info('Relatório diário concluído, inciando processamento')
            caminho_zip = resposta['Caminho ZIP']
            #Chama as macros do excel
            resposta_procss = processa_rodada_diaria(caminho_zip, pasta_relatorio)
            logging.info('Relatório diário processado')
            print('Relatório diário processado')
            break
        else:
            logging.info('Relatório ainda não concluído, inciando espera para nova verificação')
            time.sleep(4*60)
    caminho_pasta_relatorio = resposta_procss[0]
    arq_relatorio = resposta_procss[1]
    StudyId = resposta_procss[2]
    caminho_estudo_descompactado = resposta_procss[3]

    assunto = 'Boletim DIÁRIO ' + pasta_relatorio + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
    corpo = 'Prezados,\n\nSegue em anexo Boletim diário de projeção de preços (' + pasta_relatorio + ').\n\n'
    caminho_anexo = caminho_pasta_relatorio + '\\' + arq_relatorio
    anexo = arq_relatorio
    destinatario = 'maria.barbosa@eneva.com.br' #'todosenevacom@eneva.com.br' #'maria.barbosa@eneva.com.br' 
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
    logging.info('E-mail com relatório enviado, inciando arquivamento em nova base de dados')
    
    #Preenchimento da Base Prospec
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