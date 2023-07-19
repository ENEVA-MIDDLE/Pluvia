from Funcoes_API_prospec import download_estudos_finalizados_diario
from Funcoes_API_prospec import processa_rodada_diaria
import time
import logging
import datetime
from Funcoes_API_prospec import envia_email_python
from Funcoes_API_prospec import arquiva_dados_rodada_diaria

try:
    print('-----------------------------------------------------------------------')
    print('INICIANDO CÓDIGO DE MONITORAMENTO DA RODADA DIÁRIA')

    now = datetime.datetime.now()
    data_mapas = now.strftime('%d/%m/%Y')
    hora_limite = now.replace(hour=13, minute=00)

    if now < hora_limite:
        pasta_relatorio = 'Preliminar'
        x = True
    else:
        pasta_relatorio = 'Definitivo'

    data_mapas = now.strftime('%d/%m/%Y')

    nome_log = datetime.datetime.today().strftime('%Y.%m.%d') + '_log_prevs_diaria.log'
    caminho_log = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\05. Diario\01.Log'
    logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
    logging.info('--------------------------------------------------------------------------------------------------------------')
    logging.info('INICIANDO CÓDIGO DE MONITORAMENTO DA RODADA DIÁRIA')
    while True:
        logging.info('Iniciando verificação se relatório diário está concluído')
        resposta = download_estudos_finalizados_diario (pasta_relatorio)
        estudos_pendentes = resposta['Estudos pendentes']
        if estudos_pendentes == 0:
            logging.info('Relatório diário concluído, inciando processamento')
            caminho_zip = resposta['Caminho ZIP']
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
    destinatario = 'todosenevacom@eneva.com.br'
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
    logging.info('E-mail com relatório enviado, inciando arquivamento em nova base de dados')
    arquiva_dados_rodada_diaria(StudyId, data_mapas,caminho_estudo_descompactado, pasta_relatorio.upper())
    logging.info('Estudo arquivado na base de dados')
except:
    logging.exception("ERRO DE CÓDIGO")
    assunto = 'ERRO DE CÓDIGO NO BOL. DIÁRIO - ' + pasta_relatorio + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
    corpo = 'Prezados,\n\nO relatório diário (' + pasta_relatorio + ') apresentou erro no processamento. Log do erro em anexo.\n\n'
    caminho_anexo = caminho_log + '\\' + nome_log
    anexo = nome_log
    destinatario = 'maria.barbosa@eneva.com.br'
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)