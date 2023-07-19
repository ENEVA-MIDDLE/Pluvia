from Funcoes_API_prospec import autenticar_prospec
from Funcoes_API_prospec import conta_requisicoes
from Funcoes_API_prospec import runNwlistop
from Funcoes_API_prospec import getListOfDecks
from Funcoes_API_prospec import getListOfNEWAVES

token = autenticar_prospec()
print('Quantidade de requisições utilizadas: ', conta_requisicoes(token))


#oi=getListOfNEWAVES('7556')
#print(oi)

lista_decks = getListOfDecks('7556')
#print(lista_decks)
print(lista_decks[0])

#runNwlistop('7556', '165912', 'c4.large')