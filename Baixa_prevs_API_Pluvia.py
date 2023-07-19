from Funcoes_API_Pluvia import getForecasts
from Funcoes_API_Pluvia import authenticatePluvia
from Funcoes_API_Pluvia import cria_pasta_local
import datetime
from pathlib import Path
from Funcoes_API_Pluvia import downloadForecast
from Funcoes_API_Pluvia import completa_prevs
from zipfile import ZipFile
import os
import time
import pendulum

authenticatePluvia ()
#arquivo_excel = client.gencache.EnsureDispatch('Excel.Application')

ativar_completa_prevs = True
pathResult = Path(r'C:\SCRIPTS_')
    #r'C:\Users\Middle\Desktop\PREVS_PLUVIA')

#hoje  = pendulum.today(tz='America/Sao_Paulo').add(days=-2).strftime('%d/%m/%Y')

hoje = datetime.datetime.today().strftime('%d/%m/%Y')

forecastDate = hoje  #'29/09/2020'#hoje
#BIAS TRUE = SEM VIES || FALSE = COM VIES
#Opções de mapas: ONS, ECMWF_EBBCENS, ECMWF_ENS_EXT, GEFS, CFS

#ForecastModels:[{'id': 1, 'descricao': 'IA'}, {'id': 2, 'descricao': 'IA+SMAP'}, {'id': 3, 'descricao': 'SMAP'}]
#PrecipitationsDataSource:[{'id': 1, 'descricao': 'MERGE'}, {'id': 2, 'descricao': 'ETA'}, {'id': 3, 'descricao': 'GEFS'}, {'id': 4, 'descricao': 'CFS'}, {'id': 5, 'descricao': 'ONS NT0156'}, {'id': 6, 'descricao': 'Usuário'}, {'id': 7, 'descricao': 'ONS(IA+SMAP)'}, {'id': 8, 'descricao': 'Prec. Zero'}, {'id': 9, 'descricao': 'Percentis Merge'}, {'id': 10, 'descricao': 'ECMWF_ENS'}, {'id': 11, 'descricao': 'ECMWF_ENS_EXT'}, {'id': 12, 'descricao': 'ONS'}, {'id': 14, 'descricao': 'ONS_Sombra'}, {'id': 15, 'descricao': 'ONS_Pluvia'}, {'id': 16, 'descricao': 'ONS_ETAd_1_Pluvia'}]

lista_mapas = ['ONS','ECMWF_ENS','ECMWF_ENS_EXT','GEFS','CFS','PREC. ZERO']#['ONS','ECMWF_ENS', 'ECMWF_ENS_EXT', 'GEFS', 'CFS']
for mapa in lista_mapas:
    if mapa == 'ONS':
        precipitationDataSources = [12]
        forecastModels = [2]
        bias = 'True'  # True / False
        preliminary = 'False'  # True / False
        years = [2022]
        members = '' #['00', 'ENSEMBLE']
    elif mapa == 'ECMWF_ENS':
        precipitationDataSources = [10]
        forecastModels = [2]
        bias = 'False'  # True / False
        preliminary = 'False'  # True / False
        years = [2022]
        members = ['ENSEMBLE'] #, 'ENSEMBLE'
    elif mapa == 'ECMWF_ENS_EXT':
        precipitationDataSources = [11]
        forecastModels = [2]
        bias = 'False'  # True / False
        preliminary = 'False'  # True / False
        years = [2022]
        #members = ['ENSEMBLE', '00', '05', '11', '25', '41'] #, 'ENSEMBLE'
        members = ['ENSEMBLE','P50'] #, 'ENSEMBLE'['ENSEMBLE','P25','P40','P50'] 
    elif mapa == 'GEFS':
        precipitationDataSources = [3]
        forecastModels = [2]
        bias = 'False'  # True / False
        preliminary = 'False'  # True / False
        years = [2022]
        #members = ['ENSEMBLE', '00']
        members = ['ENSEMBLE']
    elif mapa == 'CFS':
        precipitationDataSources = [4]
        forecastModels = [2]
        bias = 'False'  # True / False
        preliminary = 'False'  # True / False
        years = [2022]
        #members = ['ENSEMBLE', '02', '04']
        members = ['ENSEMBLE']
    elif mapa == 'PREC. ZERO':
        precipitationDataSources = [8]
        forecastModels = [2]
        bias = 'False'  # True / False
        preliminary = 'False'  # True / False
        years = [2022]
        members = ['NULO']

    pathForecastDay = pathResult.joinpath(forecastDate[6:] + '-' + forecastDate[3:5] + '-' + forecastDate[:2])
    cria_pasta_local(pathForecastDay)
    #getFileFromAPI(token, '/api/resultados/', )
    #lista_modelos =  getInfoFromAPI(token, '/api/valoresParametros/modelos')
    #lista_mapas = getInfoFromAPI(token, '/api/valoresParametros/mapas')
    linha = 3
                            #forecastDate,precipitationDataSources,forecastModels, bias    ,preliminary, years, members
    forecasts = getForecasts(forecastDate, precipitationDataSources, forecastModels, bias, preliminary, years, members )


    #print(forecasts)
    #forecast = forecasts[0]
    for forecast in forecasts:
        nome_prevs = forecast['nome'] + ' - ' + forecast['membro'] + ' - Prevs.zip'
        downloadForecast(forecast['prevsId'], pathForecastDay, nome_prevs)
        print(forecast)
        print(forecast['nome'], ' - ', 'PrevsId: ',forecast['prevsId'])
        novo_nome_prevs = nome_prevs.replace(' - Prevs.zip','')
        with ZipFile(os.path.join(pathForecastDay, nome_prevs), 'r') as zipObj:
            zipObj.extractall(os.path.join(pathForecastDay, novo_nome_prevs))
        time.sleep(10)
        for arquivo in os.listdir(os.path.join(pathForecastDay, novo_nome_prevs)):
            novo_nome = arquivo[:12] + arquivo[-4:]
            os.rename(os.path.join(pathForecastDay, novo_nome_prevs, arquivo), os.path.join(pathForecastDay, novo_nome_prevs, novo_nome))
        if ativar_completa_prevs:
            completa_prevs(os.path.join(pathForecastDay, novo_nome_prevs))
#token1 = authenticatePluvia()
