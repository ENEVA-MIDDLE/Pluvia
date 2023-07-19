#Esse codigo apenas baixa os arquivos da rodada intersemanal e joga na base Prospec. Não há ligação com nenhum relatório final no excel.
from Funcoes_API_prospec import download_estudos_nao_finalizados_novo_arquivo, download_estudos_nao_finalizados_novo_arquivo_preco_medio_auto
from Funcoes_API_prospec import processa_rodada_intersemanal
from Funcoes_API_prospec import envia_email_python
from get_pld_semanal import get_pld_semanal
from datetime import date

today = date.today()
# today = date(2022, 8, 29)
pld_semanal = get_pld_semanal(today=today)

caminho_arquivo = r'C:\Users\alex.lourenco\OneDrive - Eneva S.A\Documentos\processos_alex\intersemanal\Criacao_Estudos_intersemanal_auto.xlsm'
#r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\04. Intersemanal\Criacao_Estudos_intersemanal_auto.xlsm'

resp_downld = download_estudos_nao_finalizados_novo_arquivo_preco_medio_auto(caminho_arquivo_estudos=caminho_arquivo, pld_semanal=pld_semanal)
caminho_rodada = resp_downld["Caminho rodada"]
print(caminho_rodada)
#processa_rodada_intersemanal(caminho_rodada)

#envia arquivo da rodada diária por e-mail para substituir a base corrente no onedrive
#caminho_pasta_base_prospec = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Prospec'
#nome_base_prospec = 'Base_prospec_diario.xlsx'
#assunto = 'BASE PROSPEC PARA ATUALIZAR NO POWER BI'# + ' - ' +  datetime.datetime.today().strftime('%d/%m/%Y')
#corpo = 'Prezados,\n\nSegue em anexo base do prospec atualizada para update do Power Bi.\n\n'
#caminho_anexo = caminho_pasta_base_prospec + '\\' + nome_base_prospec
#anexo = nome_base_prospec
#destinatario = 'alex.lourenco@eneva.com.br'
#envia_email_python(assunto, corpo, caminho_anexo, anexo, destinatario)
