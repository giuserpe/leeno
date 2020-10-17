"""
    LeenO - modulo parser XML per il formato XML-SIX
"""
import LeenoImport_XmlToscana

def parseXML(data, defaultTitle):
    '''
    estrae dal file XML i dati dell'elenco prezzi
    I dati estratti avranno il formato seguente:

        articolo = {
            'codice': codice,
            'desc': desc,
            'um': um,
            'prezzo': prezzo,
            'mdo': mdo,
            'sicurezza': oneriSic
        }
        artList = { codice : articolo, ... }

        superCatList = { codice : descrizione, ... }
        catList = { codice : descrizione, ... }

        dati = {
            'titolo': titolo,
            'superCategorie': superCatList,
            'categorie': catList,
            'articoli' : artList
        }
    '''

    # il file della Regione Sardegna è un XML-SIX con un bug
    # consistente nella mancata dichiarazione del namespace
    # quindi lo aggiungiamo a manina nei dati
    if data.find("xmlns:PRT=") < 0:
        pattern = "<PRT:Prezzario>"
        pos = data.find(pattern) + len(pattern) - 1
        data = data[:pos] + ' xmlns:PRT="mynamespace"' + data[pos:]
        print(data[:1000])

    # a parte il baco sul namespace, il formato è quello della toscana
    return LeenoImport_XmlToscana.parseXML(data, defaultTitle)

