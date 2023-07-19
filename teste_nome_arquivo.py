from win32com import client
import datetime
from Funcoes_API_Pluvia import deleta_linhas_duplicadas_data
from Funcoes_API_prospec import arquiva_estudos_intersemanal_semanal


caminho_estudo = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\2021\08\19\04.Download Estudos\E2_ECMWF_ENS_EXT-SUL-Maior-2021_9_4546.zip'
idStudy = 4546
data_criacao_estudo = '19/08/2021'
tipo_estudo = 'INTERSEMANAL'
fonte_pluvia = 'ECMWF_ENS_EXT-SUL-Maior-2021_9' #'ECMWF_ENS_EXT-SUL-Maior-2021_9'
cenario_eneva = ''
arquiva_estudos_intersemanal_semanal(idStudy, data_criacao_estudo, caminho_estudo , tipo_estudo, fonte_pluvia, cenario_eneva)
