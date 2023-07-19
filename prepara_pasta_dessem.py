from datetime import datetime
from logging import info
import pendulum
from Funcoes_API_prospec import cria_pasta_rodada_dessem
from Funcoes_API_prospec import calendario_semanas_operativas
import shutil
import os



resposta = cria_pasta_rodada_dessem(hoje=True)
data_rodada_dessem = resposta['data']
caminho_rodada_dessem = resposta['caminho']
#dia_dessem = data_rodada_dessem.format('DD')
#mes_dessem = data_rodada_dessem.format('MM')
#ano_dessem = data_rodada_dessem.format('YYYY')

def copia_arquivos_rodada_dessem(data_rodada_datetime, caminho_rodada_dessem):
    caminho_raiz_deck_dessem = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\16.DECKS\03.DESSEM'
    caminho_raiz_dadvaz = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\10.VAZÕES DIÁRIAS\04.DADVAZ'
    caminho_raiz_carga_dessem = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\01.CARGA\06.Carga Dessem'
    caminho_rodada_dessem = os.path.realpath(caminho_rodada_dessem)
    #busca o arquivo na pasta correspondente
    infos_sem_operativa = calendario_semanas_operativas(data_rodada_dessem.format('DD-MM-YYYY'))
    mes_ano_dessem = infos_sem_operativa['mes-operativo'].format('MM') + str(infos_sem_operativa['mes-operativo'].year)
    revisao = infos_sem_operativa['revisao']
    dia_dessem = infos_sem_operativa['data-referencia'].format('DD')
    mes_dessem = infos_sem_operativa['data-referencia'].format('MM')
    ano_dessem = infos_sem_operativa['data-referencia'].format('YYYY')

    caminho_dia = ano_dessem + '\\' + mes_dessem

    #copia deck dessem
    nome_arquivo_dessem = 'DS_CCEE_' + mes_ano_dessem + '_SEMREDE_' + revisao + 'D' + dia_dessem + '.zip'
    caminho_rede_deck_dessem = caminho_raiz_deck_dessem + '\\' + caminho_dia + '\\' + nome_arquivo_dessem
    shutil.copy2(caminho_rede_deck_dessem, caminho_rodada_dessem + '\\' + '01.Decks')

    #copia dadvaz para a pasta da rodada já renomeando o arquivo para dadvaz.dat
    nome_arquivo_dadvaz_final = infos_sem_operativa['data-referencia'].format('DD_MM_YYYY') + '.DAT'
    caminho_rede_dadvaz = caminho_raiz_dadvaz + '\\' + ano_dessem + '\\' + mes_dessem
    for file in os.listdir(caminho_rede_dadvaz):
        if file.endswith(nome_arquivo_dadvaz_final):
            nome_arquivo_dadvaz = file
            shutil.copy2(caminho_rede_dadvaz + '\\' + nome_arquivo_dadvaz, caminho_rodada_dessem + '\\' + '02.Prevs' + '\\' + 'dadvaz.dat')

    #copia carga dessem
    pasta_carga_rede = 'Blocos_' + infos_sem_operativa['data-referencia'].format('YYYY-MM-DD')
    caminho_pasta_carga_rede = caminho_raiz_carga_dessem + '\\' + caminho_dia + '\\' + pasta_carga_rede
    for file in os.listdir(caminho_pasta_carga_rede):
        if file == 'DE.txt' or file == 'DP.txt':
            shutil.copy2(caminho_pasta_carga_rede + '\\' + file, caminho_rodada_dessem + '\\' + '02.Prevs' + '\\' + file)
            print('Arquivo movido para a pasta da rodada:', caminho_pasta_carga_rede + '\\' + file)
