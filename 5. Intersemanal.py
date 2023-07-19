from pathlib import Path
from Funcoes_API_Pluvia import le_configuracoes_ena
from Funcoes_API_Pluvia import authenticatePluvia
from Funcoes_API_Pluvia import getForecasts
from Funcoes_API_Pluvia import cria_pasta_local
from Funcoes_API_Pluvia import downloadForecast
import warnings


warnings.filterwarnings('ignore') # ignora avisos

pathResult = Path(r'C:\Users\fernando.fidalgo\OneDrive - Eneva S.A\03. Eneva\14. Comercializadora\05. Update_ONS\01.Pluvia') #caminho de download local dos arquivos

download_arquivo = False
#traz a relação de ENAS a serem baixadas
lista_mapas = le_configuracoes_ena()

if lista_mapas == []:
    print('Nenhum mapa foi definido para download')
    quit()
authenticatePluvia ()

#este looping faz o download de cada uma das previsões da lista de mapas
for previsao in lista_mapas:
    forecasts = getForecasts(previsao['forecastDate'], previsao['mapa'], previsao['modelo'], previsao['bias'], previsao['preliminary'], previsao['years'], previsao['members'])
    if forecasts == []:
        print('Sem previsões para o Mapa ', previsao['mapa'])
    else:
        #looping em cada membro dentro de uma previsão específica
        for forecast in forecasts:
            #armazena variáveis de tempo para criar nome de pasta e nome do arquivo
            ano = previsao['forecastDate'][6:]
            mes = previsao['forecastDate'][3:5]
            dia = previsao['forecastDate'][:2]
            #transforma o valor da variável como o pluvia enxerga (false/true) para definitiva/preliminar para ser usada no nome do arquivo
            if forecast['preliminar'] == False:
                prelim = 'definitiva'
            elif forecast['preliminar'] == True:
                prelim = 'preliminar'
            else:
                prelim = 'ERRO_PRELIM'
                print('erro de código')
            #define o nome do arquivo de ENA que será baixado
            nome_mapa = forecast['nome'].replace('-' + forecast['membro'], '')
            nome_ENA = nome_mapa + '-' + forecast['membro'] + '-' + prelim + '-' + ano + mes + dia + '-ENA.zip'
            #define o caminho da pasta local para onde o arquivo de ENA será baixado
            pathForecastDay = pathResult.joinpath(ano + '-' + mes + '-' + dia)
            #cria pasta local baseada nas variáveis de data da previsão
            cria_pasta_local(pathForecastDay)
            #verifica se a ENA está disponível, caso disponível, ele faz o download daquela ENA
            if forecast['enaDisponivel']:
                downloadForecast(forecast['enaId'], pathForecastDay, nome_ENA)
                #muda a variável para True, isso indicará mais na frente que algum arquivo de ENA foi baixado dentro de todas as previsões, para que o código possa fazer o tratamento dessas informações
                download_arquivo = True
            else:
                print('ENA não disponível para o seguinte forecast-->', 'Mapa:', forecast['mapa'], ') - Modelo:',
                      forecast['modelo'])