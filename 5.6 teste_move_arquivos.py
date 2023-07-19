import os
import shutil
from pathlib import Path
novo_nome_prevs = 'CFSv2-IA+SMAP - ENSEMBLE'
pathForecastDay = Path(r'C:\Users\fernando.fidalgo\OneDrive - Eneva S.A\03. Eneva\14. Comercializadora\05. Update_ONS\01.Pluvia\2020-11-23')
caminho_rodada = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\13.RODADAS\2020\11\01\02.Prevs'
for file in os.listdir(os.path.join(pathForecastDay, novo_nome_prevs)):
    shutil.move(os.path.join(pathForecastDay, novo_nome_prevs, file), caminho_rodada + '\\' + 'CFS')
