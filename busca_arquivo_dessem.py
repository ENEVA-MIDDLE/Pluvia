from os import walk
from zipfile import ZipFile
caminho_download = r'C:\Users\fernando.fidalgo\Desktop\arquivos_download\temp'
for (dirpath, dirnames, filenames) in walk(caminho_download):
    print(filenames[0])
    deck_dessem = filenames[0]

caminho_dessem_zip = caminho_download + '\\' + deck_dessem


with ZipFile(caminho_dessem_zip, 'r') as zipObj:
        #Extract all the contents of zip file in current directory and with the same name
        #caminho da pasta é  mesmo do zip, só que sem o ".zip"
        caminho_dessem_pasta = caminho_dessem_zip.replace('.zip','') 
        zipObj.extractall(caminho_dessem_pasta)


deck_dessem_d_1 = 'DS_CCEE_072021_SEMREDE_RV3D20.zip'
for (dirpath, dirnames, filenames) in walk(caminho_dessem_pasta):
    print(filenames)

#busca arquivos que comecem com "DS"
#encontra o primeiro arquivo que começa com DS
#encontra o ultimo arquivo que começa com DS
#busca o arquivo do dia seguinte do dessem e compara com o último arquivo que começa com DS