import pendulum
from Funcoes_API_prospec import autenticar_prospec
from Funcoes_API_prospec import sendFileToDeck
from Funcoes_API_prospec import busca_deck_dessem
from win32com import client
import time
import datetime
from Funcoes_API_prospec import envia_email_python

#envia arquivo da rodada diária por e-mail para substituir a base corrente no onedrive
caminho_pasta_base_prospec = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Prospec'
nome_base_prospec = 'Base_prospec_diario.xlsx'
assunto = 'BASE PROSPEC PARA ATUALIZAR NO POWER BI'# + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
corpo = 'Prezados,\n\nSegue em anexo base do prospec atualizada para update do Power Bi.\n\n'
caminho_anexo = caminho_pasta_base_prospec + '\\' + nome_base_prospec
anexo = nome_base_prospec
destinatario = 'maria.barbosa@eneva.com.br' 
envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
#logging.info('E-mail com base de dados do prospec enviado para ativação do fluxo do Power Automate')