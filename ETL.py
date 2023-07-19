import pandas as pd
import os
from datetime import datetime
from Modules.Writer import write_to_database
#caminho_pasta=r'C:\SCRIPTS_\Temp_158\ECMWF_ENS_ext-SMAP-ENSEMBLE-definitiva-20230710-ENA'
# database = os.environ.get('SQLSERVER_DATABASE')
# table = os.environ.get('SQLSERVER_TABLE')
def to_database(caminho_pasta, database, table, mapa):
  #  ''' Exportar o Dataframe para o banco - Extract, Transforma and Load'''
    dia = int(caminho_pasta[-6:-4])
    mes = int(caminho_pasta[-8:][:-6])
    ano = int(caminho_pasta[-12:][:-8])
    data_str = datetime(ano, mes, dia).strftime('%d/%m/%Y')
    data = datetime.strptime(data_str, "%d/%m/%Y").date()
    # print(data)
    mapa=str(mapa)
        #str(caminho_pasta[-51:][:-41])
    tipo=str(caminho_pasta[-23:][:-13])
    modelo='SMAP'
        #str(caminho_pasta[-29:][:-25])
    for arquivo in os.listdir(caminho_pasta):
        if arquivo.upper().endswith('-ENA.CSV'):
          with open(caminho_pasta + '\\' + arquivo) as csv_file:
                df = pd.read_csv(csv_file, delimiter=';', skiprows=59, nrows=1715)
                df['Data Previsão']=data
                df['Mapa']=mapa
                df['Tipo de previsão']=tipo
                df['Modelo']=modelo
                df['Membro']='ENSEMBLE'
                write_to_database(data=df, database=database, table=table)

# for arquivo in os.listdir(caminho_pasta):
#         if arquivo.upper().endswith('-ENA.CSV'):
#           with open(caminho_pasta + '\\' + arquivo) as csv_file:
#                 df = pd.read_csv(csv_file, delimiter=';', skiprows=59, nrows=1715)
#                 print(df[(df['Deck']=='DC202307-sem2')& (df['Nome']=='SUDESTE') & (df['Tipo']=='Submercado')])