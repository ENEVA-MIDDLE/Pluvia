import datetime
from Funcoes_API_Pluvia import authenticatePluvia
from Funcoes_API_Pluvia import downloadForecast
from Funcoes_API_Pluvia import cria_pasta_local
from Funcoes_API_Pluvia import getForecasts
from zipfile import ZipFile
import os
import time
from pathlib import Path

def baixa_prevs_padrao (lista_mapas, **kwargs):
    #pathResult = Path(r'C:\Users\fernando.fidalgo\OneDrive - Eneva S.A\03. Eneva\14. Comercializadora\05. Update_ONS\01.Pluvia') #caminho de download
    pathResult = Path(r'C:\Users\fernando.fidalgo\pasta_pluvia') #caminho de download

    if kwargs.get('membros_EC_EXT'):
        membros_EC_EXT = kwargs.get('membros_EC_EXT')
        lista_mapas.append('ECMWF_ENS_EXT')
    authenticatePluvia()
    # arquivo_excel = client.gencache.EnsureDispatch('Excel.Application')
    hoje = datetime.datetime.today().strftime('%d/%m/%Y')
    forecastDate = '23/11/2020' #hoje  # '29/09/2020'#hoje

    for mapa in lista_mapas:
        if mapa[1] == 'ONS':
            precipitationDataSources = [12]
            forecastModels = [2]
            bias = 'True'  # True (sem viés) / False (com viés)
            preliminary = 'False'  # True(preliminar) / False (definitivo)
            years = [2020]
            members = '' #['00', 'ENSEMBLE']
        elif mapa[1] == 'ECMWF_ENS':
            precipitationDataSources = [10]
            forecastModels = [2]
            bias = 'False'  # True / False
            preliminary = 'False'  # True / False
            years = [2020]
            members = ['ENSEMBLE'] #, 'ENSEMBLE'
        elif mapa[1] == 'ECMWF_ENS_EXT':
            precipitationDataSources = [11]
            forecastModels = [2]
            bias = 'False'  # True / False
            preliminary = 'False'  # True / False
            years = [2020]
            members = membros_EC_EXT
        elif mapa[1] == 'ECMWF_ENS_EXT-00':
            precipitationDataSources = [11]
            forecastModels = [2]
            bias = 'False'  # True / False
            preliminary = 'False'  # True / False
            years = [2020]
            members = ['00'] #, 'ENSEMBLE'
        elif mapa[1] == 'ECMWF_ENS_EXT-ENSEMBLE':
            precipitationDataSources = [11]
            forecastModels = [2]
            bias = 'False'  # True / False
            preliminary = 'False'  # True / False
            years = [2020]
            members = ['ENSEMBLE']
        elif mapa[1] == 'GEFS':
            precipitationDataSources = [3]
            forecastModels = [2]
            bias = 'False'  # True / False
            preliminary = 'False'  # True / False
            years = [2020]
            members = ['ENSEMBLE']
        elif mapa == 'CFS':
            precipitationDataSources = [4]
            forecastModels = [2]
            bias = 'False'  # True / False
            preliminary = 'False'  # True / False
            years = [2020]
            members = ['ENSEMBLE']


        pathForecastDay = pathResult.joinpath(forecastDate[6:] + '-' + forecastDate[3:5] + '-' + forecastDate[:2])
        cria_pasta_local(pathForecastDay)
        # getFileFromAPI(token, '/api/resultados/', )
        # lista_modelos =  getInfoFromAPI(token, '/api/valoresParametros/modelos')
        # lista_mapas = getInfoFromAPI(token, '/api/valoresParametros/mapas')
        linha = 3
        # forecastDate,precipitationDataSources,forecastModels, bias    ,preliminary, years, members
        forecasts = getForecasts(forecastDate, precipitationDataSources, forecastModels, bias, preliminary, years,
                                 members)

        # print(forecasts)
        # forecast = forecasts[0]
        for forecast in forecasts:
            nome_prevs = forecast['nome'] + ' - ' + forecast['membro'] + ' - Prevs.zip'
            downloadForecast(forecast['prevsId'], pathForecastDay, nome_prevs)
            print(forecast)
            print(forecast['nome'], ' - ', 'PrevsId: ', forecast['prevsId'])
            novo_nome_prevs = nome_prevs.replace(' - Prevs.zip', '')
            with ZipFile(os.path.join(pathForecastDay, nome_prevs), 'r') as zipObj:
                zipObj.extractall(os.path.join(pathForecastDay, novo_nome_prevs))
            time.sleep(10)
            for arquivo in os.listdir(os.path.join(pathForecastDay, novo_nome_prevs)):
                novo_nome = arquivo[:12] + arquivo[-4:]
                os.rename(os.path.join(pathForecastDay, novo_nome_prevs, arquivo),
                          os.path.join(pathForecastDay, novo_nome_prevs, novo_nome))