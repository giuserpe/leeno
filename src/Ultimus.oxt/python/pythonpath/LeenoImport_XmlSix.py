########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

"""
    LeenO - modulo parser XML per il formato XML-SIX
"""
import re
import Dialogs
import pyleeno as PL
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
    #ripulisce il testo da caratteri non stampabili
    # data = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', data)
    data = PL.clean_text(data)

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    root = LeenoImport.stripXMLNamespaces(data)

    prezzario = root.find('prezzario')
    descrizioni = prezzario.findall('przDescrizione')
    quotazioni = prezzario.findall('listaQuotazione')
    lingue = {}
    lingua = None
    lingueEstese = {'it': 'Italiano', 'de': 'Deutsch', 'en': 'English', 'fr': 'Français', 'es': 'Español'}

    for desc in descrizioni:
        l = desc.attrib.get('lingua')
        if l:
            lExt = lingueEstese.get(l, l)
            lingue[lExt] = l
        if 'breve' in desc.attrib:
            defaultTitle = desc.attrib['breve']

    # Se non trovate o per completare, rileva le lingue dai singoli prodotti
    for product in prezzario.findall('prodotto'):
        for pDesc in product.findall('prdDescrizione'):
            l = pDesc.attrib.get('lingua')
            if l and l not in lingue.values():
                lExt = lingueEstese.get(l, l)
                lingue[lExt] = l

    listaQuotazioneId = None
    przId = prezzario.attrib.get('przId', '')
    matchAnno = re.search(r'\b(19\d\d|20\d\d)\b', przId) or re.search(r'(19\d\d|20\d\d)', przId) or re.search(r'\b(19\d\d|20\d\d)\b', defaultTitle)
    anno = matchAnno.group(1) if matchAnno else None
    if anno:
        for quota in quotazioni:
            if quota.attrib.get('lqtId') == anno:
                listaQuotazioneId = quota.attrib.get('listaQuotazioneId')
                break
    if not listaQuotazioneId and quotazioni:
        listaQuotazioneId = quotazioni[0].attrib.get('listaQuotazioneId')

    if len(lingue) > 1:
        bottoniLingue = {}
        for lName in sorted(lingue.keys(), key=lambda x: (x != 'Italiano', x)):
            bottoniLingue[lName] = lingue[lName]
        bottoniLingue['Tutte'] = 'tutte'
        bottoniLingue['Annulla'] = 'annulla'
        lingua = Dialogs.MultiButton(
            Icon="Icons-Big/question.png",
            Title="Scelta lingue",
            Text="Il file fornito è un prezzario multilinguale\n\nSelezionare la lingua da importare\noppure 'Tutte' per ottenere un prezzario multilinguale",
            Buttons=bottoniLingue)
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
    # se richiesta una lingua specifica, estrae quella
    # altrimenti le estrae tutte e le concatena una dopo l'altra
    nome = ""
    if lingua is None:
        if descrizioni:
            nome = descrizioni[0].attrib.get('breve', '')
            for desc in range(1, len(descrizioni)):
                nome = nome + '\n' + descrizioni[desc].attrib.get('breve', '')
        else:
            nome = defaultTitle
    else:
        for desc in descrizioni:
            if desc.attrib.get('lingua') == lingua:
                nome = desc.attrib.get('breve', '')
                break
        if not nome:
            nome = defaultTitle

    if nome:
        defaultTitle = nome

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
    baseText = {}
    artList = {}

    productList = prezzario.findall('prodotto')
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
        prezzo = ''
        if listaQuotazioneId:
            for el in product.findall('prdQuotazione'):
                if el.attrib.get('listaQuotazioneId') == listaQuotazioneId:
                    try:
                        prezzo = float(el.attrib['valore'])
                        break
                    except (ValueError, KeyError):
                        pass
        if prezzo == '':
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
            oneriSic = float(attr['onereSicurezza']) * prezzo / 100
        except Exception:
            oneriSic = ""
        if oneriSic == 0:
            oneriSic = ""

        # per le descrizioni: estrazione per lingua
        descs = product.findall('prdDescrizione')
        prodTextByLang = {}
        for descEl in descs:
            dAttr = descEl.attrib
            dLang = dAttr.get('lingua', '')
            breve = dAttr.get('breve', '').strip()
            estesa = dAttr.get('estesa', '').strip()
            if breve and estesa:
                if breve == estesa:
                    t = breve
                elif breve in estesa:
                    t = estesa
                elif estesa in breve:
                    t = breve
                else:
                    t = f"{breve}\n{estesa}"
            elif estesa:
                t = estesa
            elif breve:
                t = breve
            else:
                t = ""
            if t:
                if dLang in prodTextByLang:
                    prodTextByLang[dLang] += '\n' + t
                else:
                    prodTextByLang[dLang] = t

        # Gestione gerarchia: voce base e sottovoci
        base = (prezzo == "")
        if not base and (not baseCodice or not codice.startswith(baseCodice)):
            baseText = {}
            baseCodice = ''

        if base:
            baseCodice = codice
            baseText = dict(prodTextByLang)

        fullTextByLang = {}
        allLangs = set(prodTextByLang.keys())
        if not base and baseCodice and codice.startswith(baseCodice):
            allLangs |= set(baseText.keys())

        for l in allLangs:
            tBase = baseText.get(l, '') if (not base and baseCodice and codice.startswith(baseCodice)) else ''
            tChild = prodTextByLang.get(l, '')
            if tBase and tChild:
                fullTextByLang[l] = f"{tBase}\n{tChild}"
            elif tChild:
                fullTextByLang[l] = tChild
            elif tBase:
                fullTextByLang[l] = tBase

        if lingua is not None:
            # Lingua specifica (es. 'de' o 'it')
            desc = fullTextByLang.get(lingua, '')
            if not desc and '' in fullTextByLang:
                desc = fullTextByLang['']
            if not desc and fullTextByLang:
                desc = next(iter(fullTextByLang.values()))
        else:
            # Opzione 'Tutte': italiano + a capo + § + a capo + tedesco
            parti = []
            if fullTextByLang.get('it', '').strip():
                parti.append(fullTextByLang['it'].strip())
            if fullTextByLang.get('de', '').strip():
                parti.append(fullTextByLang['de'].strip())
            for lKey, lVal in fullTextByLang.items():
                if lKey not in ('it', 'de') and lVal.strip():
                    parti.append(lVal.strip())
            desc = '\n§\n'.join(parti)

        # sistema_cose
        desc = PL.clean_text(desc)

        # gruppo, nel caso ci sia
        try:
            grpId = product.find('prdGrpValore').attrib['grpValoreId']
        except Exception:
            grpId = ""

        # compone l'articolo e lo mette in lista
        # esclude dall'elenco le voci senza prezzo
        # ~if len(codice.split('.')) > 2 and prezzo != '':
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
