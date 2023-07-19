from selenium import webdriver
from Funcoes_API_Pluvia import cria_pasta_local_temporaria
import time
from Funcoes_API_prospec import calendario_semanas_operativas
from Funcoes_API_prospec import arquiva_deck_dessem
import pendulum
import os
import shutil
import datetime
import logging
from Funcoes_API_prospec import envia_email_python
from Funcoes_API_prospec import chromedriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#link = 'https://www.ccee.org.br/documents/80415/919476/DES_202111.zip/40bb512c-9818-3e40-baf6-785a1d12ea36'
#Site Novo CCEE
link = 'https://www.ccee.org.br/acervo-ccee?especie=44884&periodo=365'


endereco_download = cria_pasta_local_temporaria() #r'C:\Users\fernando.fidalgo\Desktop\Docs_Fidalgo\11.Downloads\04.CCEE'

nome_log = datetime.datetime.today().strftime('%Y.%m') + '_log_deck_dessem.log'
caminho_log = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\16.DECKS\03.DESSEM\01.Log'

logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')
logging.info('Iniciando download do deck diário do Dessem')

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
            "download.default_directory": endereco_download,  # IMPORTANT - ENDING SLASH V IMPORTANT
            "directory_upgrade": True,
            'excludeSwitches': ['enable-logging']}

try:
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option("prefs", prefs)
    driver_dessem = webdriver.Chrome(executable_path=chromedriver, chrome_options=options)
    driver_dessem.get(link)
    time.sleep(10)
    logging.info('Browser do google chrome aberto')
    infos_sem_op = calendario_semanas_operativas(hoje=True)    #(hoje=True)

    mes_dessem = infos_sem_op['mes-operativo'].format('MM') + '/' + str(infos_sem_op['mes-operativo'].format('YYYY'))
    data_dessem = infos_sem_op['data-referencia']
    
    #driver_dessem.find_element_by_xpath('//div[@id = "docsBibVirt_filter"]//input[@type = "search"]').send_keys('Dessem - ' + mes_dessem)
    #driver_dessem.find_element_by_xpath('//table[@id = "docsBibVirt"]//tbody/tr/td[@class = "sorting_1"]').click()
    #driver_dessem.find_element_by_xpath('//table[@id = "docsBibVirt"]//tbody/tr//span/a').click()
    
    #Site Novo CCEE
    #aceita os cookies
    driver_dessem.find_element_by_xpath('//div[@class = "box-cookie__buttons"]').click()
    # Busca por deck dessem
    #driver_dessem.find_element_by_xpath('//div[@class = "input-group flex-nowrap"]//input[@id = "keyword"]').send_keys('Deck de Preços - Dessem')
    driver_dessem.find_element_by_xpath('//div[@class = "input-group flex-nowrap"]//input[@id = "keyword"]').send_keys('Dessem - ' + mes_dessem)  
    # Clica no botão 'Filtrar'
    driver_dessem.find_element_by_xpath('//div[@class = "clear-filter"]//button[1]').click()
    #baixa primeiro resultado da página (assumindo que o primeiro é o mais recente)    
    driver_dessem.find_element_by_xpath('//a[@class = "d-flex ms-2 card-link"][1]').click()

    logging.info('Iniciado o download do deck do dessem')

    #Este loop aguarda até que o download do arquivo seja realiado por completo
    x = True
    while x:
        time.sleep(5)
        for file in os.listdir(endereco_download):
            if file.endswith('.zip'):
                arquivo_dessem = file
                print('Arquivo Baixado:', arquivo_dessem)
                logging.info('Download realizado:' + arquivo_dessem)
                time.sleep(5)
                driver_dessem.quit()
                x = False
                break
    #função que arquiva deck do dessem na rede
    arquiva_deck_dessem(endereco_download + '\\' + arquivo_dessem, data_dessem_datetime = data_dessem)
    logging.info('Deck dessem salvo na rede')
    shutil.rmtree(endereco_download)
    logging.info('Pasta temporária deletada')
except:
    logging.exception("ERRO DE CÓDIGO")
    assunto = 'ERRO DE CÓDIGO NO DOWNLOAD DO DECK DESSEM - '+  datetime.datetime.today().strftime('%d/%m/%Y')
    corpo = 'Prezados,\n\nO download do deck do dessem apresentou erro em sua execução. Log do erro em anexo.\n\n'
    caminho_anexo = caminho_log + '\\' + nome_log
    anexo = nome_log
    destinatario = 'maria.barbosa@eneva.com.br'
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
