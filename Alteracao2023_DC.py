from Funcoes_API_prospec import autenticar_prospec
from Funcoes_API_prospec import conta_requisicoes
from Funcoes_API_prospec import runNwlistop
from Funcoes_API_prospec import getListOfDecks
from Funcoes_API_prospec import downloadFileFromDeck
from Funcoes_API_prospec import getInfoFromStudy
from Funcoes_API_prospec import downloadDecksOfStudy
from Funcoes_API_prospec import sendFileToDeck
from Funcoes_API_prospec import executa_estudos
from Funcoes_API_prospec import escolher_sevidor
import shutil
import time
import os
import re
import pandas as pd
################################################################################################
#INPUT
################################################################################################
# ListprospecStudyId=['10725']
ListprospecStudyId=['10743','10744','10745','10746','10747','10748','10749','10750','10751','10752','10753','10754','10755','10756']
caminho_rodada=r'J:/SEDE/Comercializadora de Energia/6. MIDDLE/13.RODADAS/2022/09/02/'

################################################################################################
#Prospec
token = autenticar_prospec()
print('Quantidade de requisições utilizadas: ', conta_requisicoes(token))

rgx_mmgd = r'PQ\s+[A-Z]{3}_MMGD_[A-Z]{1}\s1O'

downloadDeckEntradaFile=True
alterarDADGER2023=True
modificar_MMGD=True
subirDADGERalterado=True
executa_rodada=True

if downloadDeckEntradaFile:
    for prospecStudyId in ListprospecStudyId:
        print(prospecStudyId)
        pathToDownload = caminho_rodada  + '06.DeckEntrada/'
        fileNameDownload = prospecStudyId + '.zip'
        #baixa deck de entrada
        downloadDecksOfStudy(prospecStudyId, pathToDownload, fileNameDownload)
        #unzip
        shutil.unpack_archive(pathToDownload +'\\' + fileNameDownload, pathToDownload +'\\' + prospecStudyId ,'zip')
        #remove pasta zipada
        os.remove(os.path.join(pathToDownload,fileNameDownload))
        decksnecessarios=['DC202301-sem1','DC202302-sem1','DC202303-sem1','DC202304-sem1','DC202305-sem1','DC202306-sem1','DC202307-sem1','DC202308-sem1','DC202309-sem1','DC202310-sem1','DC202311-sem1','DC202312-sem1']
        arquivosnecessarios=['dadger.rv0']
        #deleta pastas desnecessarias
        for file in os.listdir(pathToDownload + prospecStudyId):
            if file not in decksnecessarios:
                shutil.rmtree(pathToDownload  + prospecStudyId +'\\' + file)
        #deleta arquivos desnecessarios
        for file1 in os.listdir(pathToDownload + prospecStudyId):
            for file2 in os.listdir(pathToDownload + prospecStudyId +'\\' + file1):
                if file2 not in arquivosnecessarios:
                    os.remove(pathToDownload  + prospecStudyId +'\\' + file1 +'\\'+ file2)

        #Altera os dadger.dat -> acrescenta linhas PARPA 2023            
        if alterarDADGER2023:
            for file1 in os.listdir(pathToDownload + prospecStudyId):
                print(file1)
                for file2 in os.listdir(pathToDownload + prospecStudyId +'\\' + file1):
                    print(file2)
                    if file2 =="dadger.rv0":
                        with open(pathToDownload  + prospecStudyId +'\\' + file1 +'\\'+ file2, "a") as arquivo:
                            arquivo.write('& ESTUDO PROSPECTIVO DECOMP         => 0000 prospec.dat\n')
                            arquivo.write('& IMPRIME PREVS                     => 0001 prevs12.rv0\n')  
                            arquivo.write('& IMPRIME VAZPAST                   => 0001 vazpast2.dat\n')
                            arquivo.write('& TENDENCIA HIDROLOGICA P/ VAZPAST  => 0000\n')
                            arquivo.write('& MODELO PAR-A                      => 0001\n')
                            arquivo.close()

        # Modifica o bloco de MMGD (retirar inconsistencia "1O" -> "1")
        if modificar_MMGD:
            for file1 in os.listdir(pathToDownload + prospecStudyId):
                print(file1)
                for file2 in os.listdir(pathToDownload + prospecStudyId +'\\' + file1):
                    print(file2)
                    if file2 =="dadger.rv0":
                        with open(pathToDownload  + prospecStudyId +'\\' + file1 +'\\'+ file2, "r") as arquivo:
                            arq_str = arquivo.read()
                            all_modify = re.findall(pattern=rgx_mmgd, string=arq_str)
                            for modify in all_modify:
                                arq_str = arq_str.replace(modify, modify[:-1]+' ')
                            print(arq_str, file=open(pathToDownload  + prospecStudyId +'\\' + file1 +'\\'+ file2, 'w'))

        if subirDADGERalterado:
            #lista DECKS
            prospecStudy = getInfoFromStudy(prospecStudyId)
            listOfDecks = prospecStudy['Decks']
            print(listOfDecks)
            #envia arquivo
            for file1 in os.listdir(pathToDownload + prospecStudyId):
                print(file1)
                for file2 in os.listdir(pathToDownload + prospecStudyId +'\\' + file1):
                    print(file2)
                    if file2 =="dadger.rv0": #se o arquivo for dadger.rv0
                        nomearquivo=file1+".zip" # usa nome da pasta para identificar para qual deck vai subir DCaaaamm-sem1.zip
                        for deck in listOfDecks:
                            if deck['FileName'] == nomearquivo:
                                print(deck['Id'])
                                pathsubir=pathToDownload  + prospecStudyId +'\\' + file1 +'\\'+ file2
                                sendFileToDeck(prospecStudyId, deck['Id'], pathsubir, file2)
                            #else:
                            #    print('ATENCAO!! DECK NAO ENCONTRADO!!!')
        time.sleep(20)
        if executa_rodada:
            servidor = escolher_sevidor(prospecStudyId)
            idservidor = servidor['idservidor']
            serverType = servidor['nome_servidor']
            idQueue = 0
            ExecutionMode = 0  # Modo de execução(integer): 0 - Modo Pdrão, 1 - Consistência, 2 - Padrão + consistência
            InfeasibilityHandling = 3  # InfeasibilityHandling(integer): 0 - Parar estudo, 1 - Tratar inviabilidades, 2 - Ignorar inviabilidades, 3 - Tratar + Ignorar inviabilidades
            InfeasibilityHandlingSensibility = 3  # InfeasibilityHandlingSensibility(integer): 0 - Parar estudo, 1 - Tratar inviabilidades, 2 - Ignorar inviabilidades, 3 - Tratar + Ignorar inviabilidades
            maxRestarts = 10
            if executa_rodada:
                executa_estudos(prospecStudyId, serverType, ExecutionMode, InfeasibilityHandling, InfeasibilityHandlingSensibility, maxRestarts)
            else:
                print('Rodada  não executada.')
            

            
