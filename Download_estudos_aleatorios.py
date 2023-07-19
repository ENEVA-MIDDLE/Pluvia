from Funcoes_API_prospec import autenticar_prospec
from Funcoes_API_prospec import download_compilado
from Funcoes_API_prospec import arquiva_dados_rodada_diaria
from zipfile import ZipFile
from pathlib import Path
import os

lista_estudos = (('4045', '30/03/2021', 'DEFINITIVO'))
for estudo in lista_estudos:

    idStudy = estudo[0]
    data = estudo[1]
    prelim = estudo[2]
    #prelim = 'DEFINITIVO'


    token = autenticar_prospec()
    ano = data[-4:]
    mes = data[3:5]
    dia = data[:2]

    pathdownload = Path(r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\2021\03\30\04.Download Estudos\Definitivo')
    nome_arquivo = ano + mes + dia + '_DIARIO_' + prelim + '_' + str(idStudy) + '.zip'


    download_compilado(token, idStudy, pathdownload, nome_arquivo)
    caminho_estudo = os.path.join(pathdownload, nome_arquivo.replace('.zip', ''))
    with ZipFile(os.path.join(pathdownload, nome_arquivo), 'r') as zipObj:
        zipObj.extractall(caminho_estudo)

    arquiva_dados_rodada_diaria(idStudy, data, caminho_estudo, prelim)