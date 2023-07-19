from Funcoes_API_Pluvia import atualiza_imprime_relatorio_previsao_ENA
from Funcoes_API_prospec import envia_email_python
import datetime

hoje = datetime.datetime.today().strftime('%d/%m/%Y')
resposta = atualiza_imprime_relatorio_previsao_ENA()

assunto = 'RELATÓRIO DE PREVISÃO DIÁRIA DE ENA - FONTE PLUVIA' + hoje
corpo_email = "Prezados,\n\nSegue anexado o Relatório contendo a previsão de ENA baseada na previsão preliminar de hoje disponibilizada pelo Pluvia."
anexo = resposta[0]
print(anexo)
caminho_anexo = resposta[1]
destinatario = 'fernando.fidalgo@eneva.com.br'
envia_email_python(assunto, corpo_email, caminho_anexo, anexo, destinatario)