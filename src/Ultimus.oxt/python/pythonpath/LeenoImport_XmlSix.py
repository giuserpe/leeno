"""
    LeenO - modulo parser XML per il formato XML-SIX
"""
import Dialogs
import LeenoImport
import LeenoDialogs as DLG

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

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    root = LeenoImport.stripXMLNamespaces(data)

    prezzario = root.find('prezzario')
    descrizioni = prezzario.findall('przDescrizione')
    lingue = {}
    lingua = None
    lingueEstese = {'it': 'Italiano', 'de': 'Deutsch', 'en': 'English', 'fr': 'Français', 'es': 'Español'}
    try:
        for desc in descrizioni:
            lingua = desc.attrib['lingua']
            lExt = lingueEstese.get(lingua, lingua)
            lingue[lExt] = lingua
    except KeyError:
        pass

    if len(lingue) > 1:
        lingue['Tutte'] = 'tutte'
        lingue['Annulla'] = 'annulla'
        lingua = Dialogs.MultiButton(
            Icon="Icons-Big/question.png",
            Title="Scelta lingue",
            Text="Il file fornito è un prezzario multilinguale\n\nSelezionare la lingua da importare\noppure 'Tutte' per ottenere un prezzario multilinguale",
            Buttons=lingue)
        # se si chiude la finestra il dialogo ritorna 'None'
        # lo consideriamo come un 'Annulla'
        if lingua is None:
            lingua = 'annulla'
        if lingua == 'tutte':
            lingua = None
    else:
        lingua = None

    if lingua == 'annulla':
        return None

    # da qui, se lingua == None importa tutte le lingue presenti
    # altrimenti solo quella specificata

    # estrae il nome
    # se richiesta un lingua specifica, estrae quella
    # altrimenti le estrea tutte e le concatena una dopo l'altra
    nome = ""
    if lingua is None:
        nome = descrizioni[0].attrib['breve']
        for desc in range(1, len(descrizioni)):
            nome = nome + '\n' + descrizioni[desc].attrib['breve']
    else:
        for desc in descrizioni:
            if desc.attrib['lingua'] == lingua:
                nome = desc.attrib['breve']
                break

    # legge le unità di misura
    # siccome ci interessano soli i simboli e non il resto
    # non serve il processing per le lingue
    units = {}
    umList = prezzario.findall('unitaDiMisura')
    for um in umList:
        attr = um.attrib
        try:
            if 'simbolo' in attr:
                sym = attr['simbolo']
            else:
                sym = attr['udmId']
            umId = attr['unitaDiMisuraId']
            units[umId] = sym
        except KeyError:
            pass

    # se ci sono le categorie SOA, estrae prima quelle
    # in versione a una o più lingue a seconda del file
    # e di come viene richiesta la cosa
    # attualmente non servono, ma non si sa mai...
    categorieSOA = {}
    catList = root.findall('categoriaSOA')
    for cat in catList:
        attr = cat.attrib
        try:
            soaId = attr['soaId']
            soaCategoria = attr['soaCategoria']
            descs = cat.findall('soaDescrizione')
            text = ""
            for desc in descs:
                descAttr = desc.attrib
                try:
                    descLingua = descAttr['lingua']
                except KeyError:
                    descLingua = None
                if lingua is None or descLingua is None or lingua == descLingua:
                    text = text + descAttr['breve'] + '\n'
            if text != "":
                text = text[: -len('\n')]

            categorieSOA[soaCategoria] = {'soaId': soaId, 'descrizione': text}
        except KeyError:
            pass

    # infine tiriamo fuori il prezzario
    # utilizziamo le voci 'true' come base per le descrizioni
    # aggiungendo quelle delle voci specializzate '.a, .b...'
    baseCodice = ''
    baseTextBreve = ''
    baseTextEstesa = ''
    artList = {}

    productList = prezzario.findall('prodotto')
    madre = ''
    for product in productList:
        attr = product.attrib

        # il codice del prodotto
        if not 'prdId' in attr:
            continue
        codice = attr['prdId']

        # se c'è, estrae l'unità di misura
        if 'unitaDiMisuraId' in attr:
            um = attr['unitaDiMisuraId']
            # converte l'unità dal codice al simbolo
            um = units.get(um, "*SCONOSCIUTA*")
        else:
            # unità non trovata - la lascia in bianco
            um = ""

        # il prezzo
        # alcune voci non hanno il campo del prezzo essendo
        # voci principali composte da sottovoci
        # le importo comunque, lasciando il valore nullo
        try:
            prezzo = float(product.find('prdQuotazione').attrib['valore'])
        except Exception:
            prezzo = ""
        if prezzo == 0:
            prezzo = ""

        # percentuale manodopera
        mdo = ""
        try:
            mdo = float(product.find('incidenzaManodopera').text) / 100
        except Exception:
            mdo = ""
        if mdo == 0:
            mdo = ""

        # oneri sicurezza
        try:
            oneriSic = float(attr['onereSicurezza'])
        except Exception:
            oneriSic = ""
        if oneriSic == 0:
            oneriSic = ""

        # per le descrizioni, come sempre... processing a seconda
        # della lingua disponibile / richiesta
        descs = product.findall('prdDescrizione')
        textBreve = ""
        textEstesa = ""
        madre = ""
        for desc in descs:
            descAttr = desc.attrib
            try:
                descLingua = descAttr['lingua']
            except KeyError:
                descLingua = None
            if lingua is None or descLingua is None or lingua == descLingua:

                if 'breve' in descAttr and 'estesa' in descAttr:
                    if descAttr['breve'] in descAttr['estesa']:
                        textBreve = descAttr['estesa'] + '\n'
                    else:
                        textEstesa = descAttr['estesa'] + '\n- ' + descAttr['breve'] + '\n'
                        madre = textEstesa[: -len('\n')]
 
                    if descAttr['breve'] == descAttr['estesa']:
                        textEstesa = madre +  descAttr['breve'] + '\n'

                if 'breve' in descAttr and not 'estesa' in descAttr:
                    if descAttr['breve'][2] in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                        textEstesa = madre + descAttr['breve'] + '\n'
                    else:
                        textEstesa = madre + descAttr['breve'] + '\n'

        textBreve = textBreve.replace('Ó', 'à').replace('Þ', 'é').replace('&#x13;','').replace('&#xD;&#xA;','').replace('&#xA;','')
        textEstesa = textEstesa.replace('Ó', 'à').replace('Þ', 'é').replace('&#x13;','').replace('&#xD;&#xA;','').replace('&#xA;','')

        if textBreve != "":
            textBreve = textBreve[: -len('\n')]
        if textEstesa != "":
            textEstesa = textEstesa[: -len('\n')]

        # controlla se la voce è una voce 'base' o una specializzazione
        # della voce base. Il campo 'voce' è totalmente inaffidabile, quindi
        # consideriamo come 'base' delle voci a valore nullo e le azzeriamo
        # ad ogni nuova base e/o cambio di numerazione
        base = (prezzo == "")
        if not base and not codice.startswith(baseCodice):
            baseTextBreve = ""
            baseTextEstesa = ""

        # se voce base, tiene buona la descrizione e la salva anche come base
        # per le prossime voci (estese). Per funzionare, questo presuppone che
        # nell' XML le voci estese seguano quella base in ordine, e che tutte le voci
        # siano correttamente etichettate come voce = 'true' se voci base
        # probabilmente si può fare di meglio...
        if base:
            baseTextBreve = textBreve + '\n'
            baseTextEstesa = textEstesa + '\n'
            baseCodice = codice

        if not base and codice.startswith(baseCodice):
            textBreve = baseTextBreve +'- '+ textBreve
            textEstesa = baseTextEstesa +'- '+ textEstesa

        # utilizza solo la descrizione lunga per LeenO
        if len(textBreve) > len(textEstesa):
            desc = textBreve
        else:
            desc = textEstesa

        if len(codice.split('.')) == 4:
            madre = desc
        if len(codice.split('.')) > 4:
            if madre not in desc:
                desc = madre + desc

        # giochino per garantire che la prima stringa abbia una lunghezza minima
        # in modo che LO formatti correttamente la cella
        # ~desc = LeenoImport.fixParagraphSize(desc)

        # gruppo, nel caso ci sia
        try:
            grpId = product.find('prdGrpValore').attrib['grpValoreId']
        except Exception:
            grpId = ""

        # compone l'articolo e lo mette in lista
        # esclude dall'elenco le voci senza prezzo
        if len(codice.split('.')) > 2 and prezzo != '':
            artList[codice] = {
                'codice': codice,
                'desc': desc,
                'um': um,
                'prezzo': prezzo,
                'mdo': mdo,
                'sicurezza': oneriSic,
                'gruppo': grpId
            }

    # in alcuni casi sono presenti i gruppi, che poi sono le nostre
    # supercategorie e categorie
    # i gruppi hanno, ovviamente, una numerazione e degli ID che non c'entrano
    # un tubo con gli articoli... ma gli articoli portano un riferimento al gruppo
    # quindi una volta letti gli articoli bisogna fare uno scan a rovescio per
    # ritrovare i codici corretti delle categorie
    gruppi = {}
    superGruppi = {}
    try:
        gruppo = root.find('gruppo')
        grpValori = gruppo.findall('grpValore')
        for grpValore in grpValori:
            continue # non capisco perché, ma senza questa riga va in errore
            grpId = grpValore.attrib['grpValoreId']
            vlrId = grpValore.attrib['vlrId']
            vlrDesc = grpValore.find('vlrDescrizione').attrib['breve']
            if '.' in vlrId:
                sgId = vlrId.split('.')[0]
                gruppi[grpId] = {'cat': vlrId, 'desc': vlrDesc, 'superGroup': sgId}
            else:
                superGruppi[vlrId] = vlrDesc
    except Exception:
        pass

    # crea le categorie e supercategoria
    # è un po' un caos, ma è l'unico modo rapido per farlo
    catList = {}
    superCatList = {}
    if len(gruppi) > 0:
        for codice, articolo in artList.items():
            try:
                splitCodice = codice.split('.')
                codiceCat = splitCodice[0] + '.' + splitCodice[1]
                codiceSuperCat = splitCodice[0]
            except:
                pass
            gruppo = articolo['gruppo']
            if gruppo is None or gruppo == '':
                continue
            groupData = gruppi[gruppo]
            if not codiceCat in catList:
                catList[codiceCat] = groupData['desc']
            if not codiceSuperCat in superCatList:
                superCatList[codiceSuperCat] = superGruppi[groupData['superGroup']]

    return {
        'titolo': defaultTitle,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
    }
