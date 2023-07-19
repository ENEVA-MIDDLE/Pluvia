from Funcoes_API_prospec import chromedriver
from Funcoes_API_Pluvia import cria_pasta_local_temporaria
import shutil
import logging
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.action_chains import ActionChains
from pathlib import Path
import urllib3
import time
import pendulum
import os
from Funcoes_API_prospec import envia_email_python

link_dadvaz = 'https://sintegre.ons.org.br/sites/9/13/82/paginas/servicos/historico-de-produtos.aspx?produto=DADVAZ%20%E2%80%93%20Arquivo%20de%20Previs%C3%A3o%20de%20Vaz%C3%B5es%20Di%C3%A1rias%20(PDP)'



#data de hoje para verifique se algo publicado no dia de hoje
hoje = pendulum.today().format('DD_MM_YYYY') #'01/10/2020' #datetime.datetime.today().strftime('%d/%m/%Y')
amanha = pendulum.tomorrow().format('DD_MM_YYYY')
ontem = pendulum.yesterday().format('DD_MM_YYYY')
#depois_de_amanha = pendulum.tomorrow().add(days=1).format('DD_MM_YYYY')


#dados de log do relatório de carga
nome_log = pendulum.today().format('YYYY.MM') + '_log_dadvaz.log'   #datetime.datetime.today().strftime('%Y.%m.%d') + '_log_dadvaz.log'
caminho_log = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\10.VAZÕES DIÁRIAS\04.DADVAZ\01.Log'
logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')
logging.info('Iniciando o processo de download do DADVAZ')

#pasta padrão de download dos arquivos
endereco_download = cria_pasta_local_temporaria() 
#caminho da pasta raiz do produto que está sendo baixado, falta incluir ano e mês do arquivo dentro da pasta raiz
caminho_raiz_dadvaz = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\10.VAZÕES DIÁRIAS\04.DADVAZ'
dados_login = {
    'login_usuario': 'fernando.fidalgo@eneva.com.br',
    'login_senha': 'eneva@3444'
}

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
         "download.default_directory": endereco_download,  # IMPORTANT - ENDING SLASH V IMPORTANT
         "directory_upgrade": True,
         'excludeSwitches': ['enable-logging']}

try:
    urllib3.disable_warnings((urllib3.exceptions.InsecureRequestWarning))
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option("prefs", prefs)
    driver_dadvaz = webdriver.Chrome(executable_path=chromedriver, chrome_options=options)
    #abre o chrome e entra no link fornecido
    driver_dadvaz.get(link_dadvaz)
    driver_dadvaz.find_element_by_id('username').send_keys(dados_login['login_usuario']) # insere dados de login
    driver_dadvaz.find_element_by_name('submit.IdentificarUsuario').click() #clica para aparecer a opção de senha
    time.sleep(5)
    driver_dadvaz.find_element_by_id('password').send_keys(dados_login['login_senha']) #insere dados de senha
    driver_dadvaz.find_element_by_name('submit.Signin').click() # clica para logar
    logging.info('Logado com sucesso no site do ONS')
    print ('Logado com sucesso!')
    time.sleep(10) #aguarda página carregar
    #clica no botão para concordar com os cookies e abrir uma janela maior do navegador
    barra_rodape = driver_dadvaz.find_element_by_id('ons_terms')
    botao = barra_rodape.find_element_by_xpath('.//button')
    botao.click()
    print('botao clicado')
    time.sleep(10)

    #move até o elemento1 e clica para download
    elemento1 = driver_dadvaz.find_element_by_xpath('//a[contains(text(),"' + hoje + '.DAT")]')
    actions = ActionChains(driver_dadvaz)
    
    actions.move_to_element(elemento1).click().perform()
    time.sleep(5)
    print('primeiro scrow')
    #move até o elemento2 e clica para download
    elemento2 = driver_dadvaz.find_element_by_xpath('//a[contains(text(),"' + amanha + '.DAT")]')
    actions = ActionChains(driver_dadvaz)
    actions.move_to_element(elemento2).click().perform()
    time.sleep(5)

    print('Achou o elemento')

    logging.info('Iniciado o download do DADVAZ (D e D+1)')
    time.sleep(10)
    driver_dadvaz.quit()
    #busca arquivos na pasta de download
    for file in os.listdir(endereco_download):
        nome_arquivo_dadvaz = file
        print('Arquivo baixado:', nome_arquivo_dadvaz)
        logging.info('Download realizado:' + nome_arquivo_dadvaz)
        ano_dadvaz = nome_arquivo_dadvaz[-8:-4]
        mes_dadvaz = nome_arquivo_dadvaz[-11:-9]
        caminho_rede_dadvaz = caminho_raiz_dadvaz + '\\' + ano_dadvaz + '\\' + mes_dadvaz
        os.makedirs(caminho_rede_dadvaz, exist_ok=True)
        shutil.move(endereco_download + '\\' + nome_arquivo_dadvaz, caminho_rede_dadvaz + '\\' + nome_arquivo_dadvaz)
        print('Arquivo movido para a rede:', nome_arquivo_dadvaz)
        logging.info('Arquivo movido para a rede:'+ nome_arquivo_dadvaz)
    shutil.rmtree(endereco_download)
    print('Pasta temporária deletada')
    logging.info('Pasta local temporária deletada')

except:
    logging.exception("ERRO DE CÓDIGO")
    assunto = 'ERRO DE CÓDIGO NO DOWNLOAD DO DADVAZ - '+  pendulum.today().format('DD/MM/YYY.MM') 
    corpo = 'Prezados,\n\nO download do dadvaz apresentou erro em sua execução. Log do erro em anexo.\n\n'
    caminho_anexo = caminho_log + '\\' + nome_log
    anexo = nome_log
    destinatario = 'maria.barbosa@eneva.com.br'
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
