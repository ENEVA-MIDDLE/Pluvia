from typing import final
from Funcoes_API_prospec import calendario_semanas_operativas
import pendulum
dia_dessem = pendulum.today().format('DD-MM-YYYY')
print('dia dessem:', dia_dessem)
resposta = calendario_semanas_operativas(dia_dessem)

print(resposta)
data_fim = resposta['final']
mes_operativo = data_fim.replace(day = 1)
inicio_mes_operativo = mes_operativo.subtract(days=(mes_operativo.day_of_week+1)) 
mes_seguinte = mes_operativo.add(months = 1)
termino_mes_operativo = mes_seguinte.subtract(days=(mes_seguinte.day_of_week+2))
print('Mês Operativo:', mes_operativo)
print('Mês Seguinte:', mes_seguinte)
print('Início mês Operativo:', inicio_mes_operativo)
print('Término mês Operativo:', termino_mes_operativo)

mes_ano_dessem = '%02d' % mes_operativo.month + str(mes_operativo.year)
#print(mes_dessem)

nome_arquivo_dessem = 'DS_CCEE_' + mes_ano_dessem + '_SEMREDE_REV' + str(resposta['rev']) + 'D' + dia_dessem[:2] + '.zip'
print(nome_arquivo_dessem)
print()
caminho_deck_rede = caminho_raiz_rede + '\\' + str(infos_sem_operativa['data-referencia'].format('YYYY')) + '\\' + str(infos_sem_operativa['data-referencia']('MM'))