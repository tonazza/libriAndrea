import xlwings
import CONFIG
import SECRETS
from icecream import ic
import requests
import json
from time import sleep

#ic.disable()

def aggiorna_lista_libri(listaLibri):
    
    prima_riga = CONFIG.PRIMA_RIGA_TABELLA
    ultima_riga = listaLibri.used_range.last_cell.row
    ic("righe da esaminare: {} - {}".format(prima_riga,ultima_riga))
    
    riga = prima_riga
    while (riga<=ultima_riga):
        autori = listaLibri.range((riga, CONFIG.COLONNA_CONTROLLO_RIGA_COMPILATA)).value
        if (autori == None):
            isbn = listaLibri.range((riga, CONFIG.COLONNA_ISBN)).value
            if (isbn != None):
                ic("Completo la riga {} ...".format(riga))
                # se c'è l'ISBN faccio la richiesta e riempio le celle
                apiUrl = SECRETS.STRINGA_SX_API + isbn + SECRETS.STRINGA_DX_API
                try:
                    rispostaGrezza = requests.get(apiUrl)
                    if (rispostaGrezza.status_code ==200):
                        risposta = rispostaGrezza.json()
                        autori = json.dumps(risposta['items'][0]['volumeInfo']['authors'])
                        titolo =risposta['items'][0]['volumeInfo']['title']
                        data = risposta['items'][0]['volumeInfo']['publishedDate']
                        ic("{} {} {}".format(autori, titolo, data))
                        listaLibri.range((riga, CONFIG.COLONNA_AUTORI)).value = autori
                        listaLibri.range((riga, CONFIG.COLONNA_TITOLO)).value = titolo
                        listaLibri.range((riga, CONFIG.COLONNA_DATA)).value = data
                    else:
                        print("Status code API: {}".format(rispostaGrezza.status_code))
                    sleep(0.5)
                except:
                    ic("errore")
        
        riga = riga+1
        


if __name__ == "__main__":
    #apertura file
    listaLibri = xlwings.Book(CONFIG.PERCORSO_FILE_EXCEL)
    foglioLibri = listaLibri.sheets(CONFIG.NOME_FOGLIO)
    
    try:
        aggiorna_lista_libri(foglioLibri)
        listaLibri.save()
        #listaLibri.close()
    
    except Exception as e:
        ic(e)
    