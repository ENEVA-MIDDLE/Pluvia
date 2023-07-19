from Funcoes_API_prospec import download_estudos_nao_finalizados_novo_arquivo
from Funcoes_API_prospec import download_estudos_nao_finalizados_novo_arquivo_semanal, download_estudos_nao_finalizados_novo_arquivo_preco_medio_auto
from Funcoes_API_prospec import processa_rodada_intersemanal
from Funcoes_API_prospec import envia_email_python
from get_pld_semanal import get_pld_semanal
from datetime import date

today = date.today()
# today = date(2022, 8, 29)
pld_semanal = get_pld_semanal(today=today)

#caminho_arquivo = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\Criacao_Estudos.xlsm'
caminho_arquivo = r'C:\SCRIPTS_\Criacao_Estudos.xlsm'

#sem calculo mensal
resp_downld = download_estudos_nao_finalizados_novo_arquivo_preco_medio_auto(pld_semanal=pld_semanal)
caminho_rodada = resp_downld["Caminho rodada"]
print(caminho_rodada)

#calculo mensal
#download_estudos_nao_finalizados_novo_arquivo_semanal()


#Macro
#processa_rodada_intersemanal(caminho_rodada)

#envia arquivo da rodada diária por e-mail para substituir a base corrente no onedrive
#caminho_pasta_base_prospec = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Prospec'
#nome_base_prospec = 'Base_prospec_diario.xlsx'
#assunto = 'BASE PROSPEC PARA ATUALIZAR NO POWER BI'# + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
#corpo = 'Prezados,\n\nSegue em anexo base do prospec atualizada para update do Power Bi.\n\n'
#caminho_anexo = caminho_pasta_base_prospec + '\\' + nome_base_prospec
#anexo = nome_base_prospec
#destinatario = 'renata.hunder@eneva.com.br; alex.lourenco@eneva.com.br'
#envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)