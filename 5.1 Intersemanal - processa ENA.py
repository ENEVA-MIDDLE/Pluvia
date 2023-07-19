from zipfile import ZipFile
import os
from Funcoes_API_Pluvia import le_ena_pasta
from Funcoes_API_Pluvia import  deleta_linhas_duplicadas
from Funcoes_API_Pluvia import salva_ENA_base
from Funcoes_API_Pluvia import copia_arquivo_base_ena
import shutil
from pathlib import Path

pathForecastDay = r'C:\Users\fernando.fidalgo\OneDrive - Eneva S.A\03. Eneva\14. Comercializadora\05. Update_ONS\01.Pluvia\2020-11-16'
data_ec = '16/11/2020'
caminho_pasta_ENA = copia_arquivo_base_ena(data_ec)
for arquivo in os.listdir(pathForecastDay):  # lê os arquivos que estão na pasta de download
    if arquivo.upper().endswith('ENA.ZIP'):  # exibe somente os arquvos de ENA
        caminho_arquivo = os.path.realpath(pathForecastDay) + '\\' + arquivo
        caminho_pasta = caminho_arquivo.replace('-ENA.zip', '')
        with ZipFile(caminho_arquivo, 'r') as zipObj:  # descompacta o zip na mesma pasta do zip
            zipObj.extractall(caminho_pasta)
            print('arquivo descompactado:', arquivo)
        # função que lê o arquivo CSV de ENA dentro da pasta e retorna uma lista onde cada item da lista é uma linha, e cada linha é uma nova lista, onde cada item é uma coluna
        tabela = []
        tabela = le_ena_pasta(caminho_pasta)
        data_mapa = tabela[0][0]
        mapa_atual = tabela[0][1]
        modelo = tabela[0][2]
        membro = tabela[0][3]
        #DECIDIR sE VOU COLOCAR ESSA FUNÇÃO DE DELETAR LINHAS, POIS NÃO TERÃO LINHAS DUPLICATAS
        #deleta linhas do arquivo de base, de acordo com data da previsão, mapa, modelo e membro da previsão
        #deleta_linhas_duplicadas(data_mapa, mapa_atual, modelo, membro)

        #--------------------------------------------------------------------------------------------------------------------------
        #INSERIR AQUI FUNCAO PARA COPIAR UM ARQUIVO BASE DE ENA PARA A PASTA DA RODADA E RETORNAR O CAMINHO DA RODADA
        # --------------------------------------------------------------------------------------------------------------------------
       # caminho_pasta_ENA = copia_arquivo_base_ena(data_ec)


        salva_ENA_base(tabela, caminho_salvar_base_pluvia = caminho_pasta_ENA)  # função que lê a tabela gerada do arquivo de ENA lido e salva na base em excel


        shutil.rmtree(caminho_pasta)
        os.remove(pathForecastDay + '\\' + arquivo)
        #print('Pasta local deletada: ', caminho_pasta)