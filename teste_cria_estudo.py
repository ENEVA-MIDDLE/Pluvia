import os
from win32com import client
from pathlib import Path
import shutil
import time
caminho_local_raiz =  r'C:\Users\fernando.fidalgo\Desktop\Docs_Fidalgo\10. Eneva Com'
caminho_base_pluvia = r'J:\SEDE\Comercializadora de Energia\6. MIDDLE\12.ANALISES\Base_ENA_Pluvia.xlsx'

caminho_base_pluvia_realpath = os.path.realpath(caminho_base_pluvia)
caracter_ena = caminho_base_pluvia_realpath.rfind('\\') + 1
nome_arquivo_ena = caminho_base_pluvia_realpath[caracter_ena:]
nome_arquivo_temp = nome_arquivo_ena[:-5] + '_temp' + '.xlsx'
caminho_arquivo_temp = os.path.join(Path(caminho_local_raiz), nome_arquivo_temp)
print('Inciando cópia de arquivo original da rede no caminho: ', caminho_base_pluvia_realpath)
shutil.copy2(caminho_base_pluvia, caminho_arquivo_temp)
print('Arquivo temporário copiado para pasta local: ', nome_arquivo_temp)
time.sleep(5)



arquivo_excel_base_pluvia = client.DispatchEx("Excel.Application")
arquivo_excel_base_pluvia.Visible = True
wb_base_ena = arquivo_excel_base_pluvia.Workbooks.Open(Filename=caminho_arquivo_temp)
ws_base_ena = wb_base_ena.Worksheets('pluvia_definitivo')
lin = ws_base_ena.UsedRange.Rows.Count + 1
print(lin)
ws_base_ena.Cells(lin, 1).Value = 'HAHAHAHA'
#wb_base_ena.Save
wb_base_ena.Close(True)
arquivo_excel_base_pluvia.Quit()
del arquivo_excel_base_pluvia
shutil.copy2(caminho_base_pluvia, caminho_arquivo_temp)

