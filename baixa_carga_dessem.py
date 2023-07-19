from Funcoes_API_prospec import chromedriver
from Funcoes_API_Pluvia import cria_pasta_local_temporaria
import shutil
import logging
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from pathlib import Path
import urllib3
import time
import pendulum
import os
from zipfile import ZipFile
from Funcoes_API_prospec import envia_email_python

#mflb
#Ajustei para baixar o arquivo do dia seguinte (considerando que o cod vai rodar 13h)
#Antes baixava o arquivo do dia pq rodava meia noite do dia seguinte a publicacao(a publicacao é sempre referente ao dia seguinte)
#elemento 1

link_carga_dessem = 'https://sintegre.ons.org.br/sites/9/46/paginas/servicos/historico-de-produtos.aspx?produto=Arquivos%20de%20Previs%C3%A3o%20de%20Carga%20para%20o%20DESSEM'



#data de hoje para verifique se algo publicado no dia de hoje
hoje = pendulum.today().format('YYYY-MM-DD') 
amanha = pendulum.tomorrow().format('YYYY-MM-DD')
#depois_de_amanha = pendulum.tomorrow().add(days=1).format('DD_MM_YYYY')


#dados de log do relatório de carga
nome_log = pendulum.today().format('YYYY.MM') + '_log_carga_dessem.log'   
caminho_log = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\01.CARGA\06.Carga Dessem\01.Log'
logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')
logging.info('Iniciando o processo de download da carga diária do Dessem')



#pasta padrão de download dos arquivos
endereco_download = cria_pasta_local_temporaria() 
#caminho da pasta raiz do produto que está sendo baixado, falta incluir ano e mês do arquivo dentro da pasta raiz
caminho_raiz_carga_dessem = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\01.CARGA\06.Carga Dessem'
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
    driver_carga_dessem = webdriver.Chrome(executable_path=chromedriver, chrome_options=options)

    #abre o chrome e entra no link fornecido

    driver_carga_dessem.get(link_carga_dessem)
    driver_carga_dessem.find_element_by_id('username').send_keys(dados_login['login_usuario']) # insere dados de login
    driver_carga_dessem.find_element_by_name('submit.IdentificarUsuario').click() #clica para aparecer a opção de senha
    time.sleep(5)
    driver_carga_dessem.find_element_by_id('password').send_keys(dados_login['login_senha']) #insere dados de senha
    driver_carga_dessem.find_element_by_name('submit.Signin').click() # clica para logar
    print ('Logado com sucesso!')
    logging.info('Logado com sucesso no site do ONS')
    time.sleep(10) #aguarda página carregar
    #busca a barra de rodapé do ONS e clica nela para a tela ficar maior
    try:

        barra_rodape = driver_carga_dessem.find_element_by_id('ons_terms')
        botao = barra_rodape.find_element_by_xpath('.//button')
        botao.click()
        time.sleep(5)
        #elemento1 = driver_carga_dessem.find_element_by_xpath('//a[contains(text(),"' + hoje + '.zip")]')
        elemento1 = driver_carga_dessem.find_element_by_xpath('//a[contains(text(),"' + amanha + '.zip")]')
        #elemento2 = driver_carga_dessem.find_element_by_xpath('//a[contains(text(),"' + amanha + '.DAT")]')
        print('Achou o elemento')
        
        actions = ActionChains(driver_carga_dessem)
        #move a página até o elemento
        actions.move_to_element(elemento1).click().perform()
        time.sleep(5)
        #elemento1.click()
        #elemento2.click()
    except:
        driver_carga_dessem.refresh()
        time.sleep(10)
        barra_rodape = driver_carga_dessem.find_element_by_id('ons_terms')
        botao = barra_rodape.find_element_by_xpath('.//button')
        botao.click()
        time.sleep(5)
        elemento1 = driver_carga_dessem.find_element_by_xpath('//a[contains(text(),"' + hoje + '.zip")]')
        #elemento2 = driver_carga_dessem.find_element_by_xpath('//a[contains(text(),"' + amanha + '.DAT")]')
        print('Achou o elemento')
        
        actions = ActionChains(driver_carga_dessem)
        #move a página até o elemento
        actions.move_to_element(elemento1).click().perform()
        time.sleep(5)
        #elemento1.click()
        #elemento2.click()

    #Este loop aguarda até que o download do arquivo seja realizado por completo
    logging.info('Iniciado o download de carga do dessem')
    x = True
    while x:
        time.sleep(5)
        for file in os.listdir(endereco_download):
            #verifica se na pasta já contem um aruqivo ZIP, caso contrário continua o loop
            if file.endswith('.zip'):
                nome_arquivo_carga_dessem = file
                print('Arquivo baixado:', nome_arquivo_carga_dessem)
                logging.info('Download realizado:' + nome_arquivo_carga_dessem)
                caminho_arquivo_carga_dessem_ZIP = endereco_download + '\\' + nome_arquivo_carga_dessem
                driver_carga_dessem.quit()
                x = False
                break

    #inicia tratativas para arquivar na rede
    #descompacta arquivo localmente
    with ZipFile(caminho_arquivo_carga_dessem_ZIP, 'r') as zipObj:
        #Extract all the contents of zip file in current directory and with the same name
        #caminho da pasta é  mesmo do zip, só que sem o ".zip"
        caminho_carga_dessem_pasta = caminho_arquivo_carga_dessem_ZIP.replace('.zip','') 
        zipObj.extractall(caminho_carga_dessem_pasta)
        print('Arquivo descompactado em pasta local temporária:', caminho_carga_dessem_pasta)
        logging.info('Arquivo descompactado em pasta local temporária:' + caminho_carga_dessem_pasta)



    ano_carga_dessem = nome_arquivo_carga_dessem[7:11]
    mes_carga_dessem = nome_arquivo_carga_dessem[12:14]
    caminho_rede_carga_dessem = caminho_raiz_carga_dessem + '\\' + ano_carga_dessem + '\\' + mes_carga_dessem
    caminho_pasta_rede = caminho_rede_carga_dessem + '\\' + nome_arquivo_carga_dessem.replace('.zip','')
    if os.path.exists(caminho_pasta_rede):
        logging.info('Pasta já existente na rede:' + caminho_pasta_rede)
        shutil.rmtree(caminho_pasta_rede)
        logging.info('Deletada pasta anterior já existente na rede:'+ caminho_pasta_rede)
    #os.makedirs(caminho_raiz_carga_dessem, exist_ok=True)
    os.makedirs(caminho_rede_carga_dessem, exist_ok=True)
    shutil.move(caminho_carga_dessem_pasta, caminho_rede_carga_dessem)
    logging.info('Arquivo movido para a rede:'+ caminho_pasta_rede)

    print('Arquivo movido para a rede:'+ caminho_rede_carga_dessem)


    shutil.rmtree(endereco_download)
    print('Pasta temporária deletada')
    logging.info('Pasta local temporária deletada')
except:
    logging.exception("ERRO DE CÓDIGO")
    assunto = 'ERRO DE CÓDIGO NO DOWNLOAD DA previsão de carga diária do dessem - '+  pendulum.today().format('DD/MM/YYY.MM') 
    corpo = 'Prezados,\n\nO download da carga diária do dessem apresentou erro em sua execução. Log do erro em anexo.\n\n'
    caminho_anexo = caminho_log + '\\' + nome_log
    anexo = nome_log
    destinatario = 'maria.barbosa@eneva.com.br'
    envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)

