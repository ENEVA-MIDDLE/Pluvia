# Esse codigo não baixa os prevs do pluvia e não faz as estatísticas
from Funcoes_API_Pluvia import baixa_ENA_configurada
from Funcoes_API_Pluvia import arquiva_base_ENA
from Funcoes_API_Pluvia import calcula_estatistica_ENA
from Funcoes_API_Pluvia import analise_resultado_estatistica
from Funcoes_API_Pluvia import compila_baixa_prevs_intersemanal
from Funcoes_API_prospec import cria_rodada
import datetime
import logging
nome_log = datetime.datetime.today().strftime('%Y.%m.%d') + '_log_intersem.log'
caminho_log = r'C:\SCRIPTS_\Log'
    #r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\04. Intersemanal\01.Log'

logging.basicConfig(filename=caminho_log + '\\' + nome_log, format='%(asctime)s - %(message)s', datefmt='%d-%m-%Y %H:%M:%S', level=logging.INFO)
logging.info('--------------------------------------------------------------------------------------------------------------')
caminho_arquivo_cria_estudo = r'C:\Users\alex.lourenco\OneDrive - Eneva S.A\Documentos\processos_alex\intersemanal\Criacao_Estudos_intersemanal_auto.xlsm'
#lê o arquivo de configuração de mapas do intersemanal (para o caso de querer incluir o CFS, ainda a configurar)
#a princípio essa função baixa apenas os membros do EC extendido, a ideia é poder utilizar o CFS no futuro também
#necessário incluir uma forma de baixar a previsão do dia que quiser, hoje está configurado para para baixar do próprio dia
data_mapas = datetime.datetime.today().strftime('%d/%m/%Y') #'30/11/2020'
#data_mapas = '20/12/2021'

#logging.info('Iniciando download das ENAs para data de referência:' + data_mapas)
#retorna o caminho local (raiz) onde os arquivos baixados foram salvos
#resposta = baixa_ENA_configurada(data_mapas) #código 5. Intersemanal.py
#caminho_local = resposta[0] #caminho local onde foram salvas
#logging.info('Caminho local onde foi feito o download das ENAs:' + str(caminho_local))
#print('caminho local:',caminho_local)
#data_mapa = resposta[1] #data em que a ENA do EC estendido foi compilada

#logging.info('Iniciando a compilação da ENA em arquivo excel na rede')
#caminho_base_ENA = arquiva_base_ENA (data_mapa, caminho_local) #código 5.1 Intersemanal - processa ENA
#print('Caminho base ENA:',caminho_base_ENA)
#logging.info('ENA arquivada no caminho:' + str(caminho_base_ENA))

#logging.info('Iniciando cálculos estatísticos na base de ENA.')
#calcula_estatistica_ENA (caminho_base_ENA) #monta a tabela com estatisticas das ENAS baixadas # 5.2 Estatistica_ENA.py
#logging.info('Realizados os cálculos estatísticos na base de ENA.')

#logging.info('Iniciando a análise dos resultados estatícticos da base de ENA')
#analise_resultado_estatistica (caminho_arquivo_cria_estudo, caminho_base_ENA) #código 5.3 Le estatistica.py
#logging.info('Análise esttaística concluída e membros identificados para download de Prevs')
#caminho_rodada = caminho_base_ENA[:caminho_base_ENA[:caminho_base_ENA.rfind('\\')].rfind('\\')]
#logging.info('Iniciando download de Prevs conforme planilha de criação de Estudos.')
#compila_baixa_prevs_intersemanal (caminho_arquivo_cria_estudo, data_mapa) #código 5.4 Lista_Prevs_rodada.py
#logging.info('Conluído o download de Prevs e feitas as devidas tratativas nos arquivos de prevs.')
#logging.info('Iniciando o upload de informações da rodada no Prospec.')

cria_rodada(caminho_arquivo_cria_estudo, Executa_rodada='True')
logging.info('Rodada criada com sucesso no prospec.')
##############################################################################################################
#FALTA BAIXAR AUTOMATICAMENTE E RODAR MACROS DE RESULTADOS E MATRIZ
#FALTRA ENVIAR RESULTADO POR E-MAIL AUTOMATICAMENTE
##############################################################################################################


