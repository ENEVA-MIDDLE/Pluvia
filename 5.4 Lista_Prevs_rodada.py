from win32com import client
from Funcoes_API_Pluvia import baixa_prevs_padrao

caminho_arquivo_cria_estudo = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\04. Intersemanal\Testes\Criacao_Estudos_intersemanal_auto_teste.xlsm'

arquivo_excel = client.DispatchEx("Excel.Application")
wb_study = arquivo_excel.Workbooks.Open(Filename = caminho_arquivo_cria_estudo)
ws_config = wb_study.Worksheets('configuracoes')
mapas_fixos = []
membros_EC_ext = []
linha = 9
if ws_config.Cells(linha, 6).Value is None:
    print('Rodada sem Mapas fixos')
else:
    while ws_config.Cells(linha, 6).Value is not None:
        mapas_fixos.append((ws_config.Cells(linha, 6).Value, ws_config.Cells(linha, 8).Value))
        linha = linha + 1
linha = 26
if ws_config.Cells(linha, 6).Value is None:
    print('Rodada sem Mapas com Membros')
else:
    while ws_config.Cells(linha, 6).Value is not None:
        membros_EC_ext.append((ws_config.Cells(linha, 13).Value, ws_config.Cells(linha, 14).Value ))
        linha = linha + 1
    linha = 26

print(mapas_fixos)
print(membros_EC_ext)
#caminho da rodada
baixa_prevs_padrao (mapas_fixos, membros_EC_EXT = membros_EC_ext)

wb_study.Close(False)
arquivo_excel.Application.Quit()