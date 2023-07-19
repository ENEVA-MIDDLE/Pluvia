from time import process_time_ns
from Funcoes_API_prospec import calendario_semanas_operativas
import pendulum

data_dessem = pendulum.datetime(2021, 7, 15, tz='America/Sao_Paulo')
data_formatada = data_dessem.format('DD-MM-YYYY')
infos_semana_operativa = calendario_semanas_operativas(data_formatada)
print(infos_semana_operativa)

final_semana_operativa = pendulum.datetime(2021, 5,1, tz='America/Sao_Paulo')
#print(final_semana_operativa.day_of_week)