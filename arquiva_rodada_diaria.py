from Funcoes_API_prospec import arquiva_dados_rodada_diaria

idStudy = 3785
data = '01/02/2021'
caminho_estudo = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\2021\02\01\04.Download Estudos\Definitivo\1.0_202121_REV0_FEV-21_A_FEV-21_DIARIO_DEFINITIVO_3785'
#prelim = 'PRELIMINAR'
prelim = 'DEFINITIVO'


arquiva_dados_rodada_diaria(idStudy, data, caminho_estudo, prelim)