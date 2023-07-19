from Funcoes_API_prospec import autenticar_prospec
from Funcoes_API_prospec import getListOfDecks
from Funcoes_API_prospec import api_url_base
from Funcoes_API_prospec import verifyCertificate
from pathlib import Path
import requests
import os
import pendulum
from Funcoes_API_prospec import download_decks_iniciais
from Funcoes_API_Pluvia import cria_pasta_local_temporaria


print(pendulum.now('America/Sao_Paulo').format('DD/MM/YYYY HH:mm:ss'))
token = autenticar_prospec()

studyid = 4512

nome_deck_entrada_zip = str(studyid) + '_decks_entrada.zip'
caminho_download = cria_pasta_local_temporaria()

download_decks_iniciais(token, studyid, caminho_download, nome_deck_entrada_zip)
