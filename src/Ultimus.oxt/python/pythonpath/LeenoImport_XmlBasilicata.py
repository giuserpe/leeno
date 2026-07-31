import re
import pyleeno as PL
import LeenoImport
import xml.etree.ElementTree as ET
# ~import LeenoDialogs as DLG

# ~from com.sun.star.sheet.CellFlags import \
    # ~VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED

def parseXML(data, defaultTitle=None):
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

    # alcuni files sono degli XML-SIX con un bug
    # consistente nella mancata dichiarazione del namespace
    # quindi lo aggiungiamo a manina nei dati

    # ~ if data.find("xmlns=") < 0:
        # ~ pattern = "<PRT:Prezzario>"
        # ~ pos = data.find(pattern) + len(pattern) - 1
        # ~ data = data[:pos] + ' xmlns:PRT="mynamespace"' + data[pos:]
        # ~ print(data[:1000])

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    root = LeenoImport.stripXMLNamespaces(data)

    # Due varianti note dello stesso schema:
    # - formato "storico": il titolo è ricavato dal tag <pdf> in radice
    #   (nome del file PDF da cui il prezzario è stato originariamente estratto)
    # - formato 2026 (es. INFRA_Prezzario_Regionale_Basilicata_2026.xml):
    #   non c'è il tag <pdf>, ma in radice è presente <anno>;
    #   il resto della struttura (capitoli/categorie/voci/sottovoci)
    #   è identico al formato storico.
    pdfNode = root.find('pdf')
    if pdfNode is not None and pdfNode.text:
        titolo = pdfNode.text
        if '.pdf' in titolo:
            titolo = titolo[: -4]

        titolo = ' '.join(titolo.split('_'))
    else:
        annoNode = root.find('anno')
        anno = annoNode.text.strip() if annoNode is not None and annoNode.text else ''

        if defaultTitle:
            titolo = defaultTitle
        elif anno:
            titolo = 'Elenco prezzi - Regione Basilicata - anno ' + anno
        else:
            titolo = 'Elenco prezzi - Regione Basilicata'

    artList = {}
    superCatList = {}
    catList = {}

    capitoli = root.find('capitoli') # è una sola ricorrenza
    if capitoli is None:
        capitoli = []

    for Capitolo in capitoli:

        # estrae supercategoria e categoria
        codiceSuperCatNode = Capitolo.find('codice')
        if codiceSuperCatNode is None or not codiceSuperCatNode.text:
            # capitolo privo di codice: non è utilizzabile, si salta
            continue
        codiceSuperCat = codiceSuperCatNode.text
        descSuperCatNode = Capitolo.find('descrizione')
        superCat = descSuperCatNode.text.strip() if descSuperCatNode is not None and descSuperCatNode.text else ''

        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        categorie = Capitolo.find('categorie')
        if categorie is None:
            # capitolo senza sotto-categorie: niente da estrarre, si passa oltre
            continue

        for Categoria in categorie:
            codiceCatNode = Categoria.find('codice')
            if codiceCatNode is None or not codiceCatNode.text:
                continue
            codiceCat = codiceSuperCat + '.' + codiceCatNode.text
            CatDescNode = Categoria.find('descrizione')
            Cat = CatDescNode.text if CatDescNode is not None else ''
            if not codiceCat in catList:
                catList[codiceCat] = Cat

            # estrae voci e sottovoci
            voci = Categoria.find('voci')
            if voci is None:
                # categoria senza voci: niente da estrarre, si passa oltre
                continue

            for Voce in voci:
                voceDescNode = Voce.find('descrizione')
                voce = voceDescNode.text if voceDescNode is not None and voceDescNode.text else ''
                voceCodiceNode = Voce.find('codice')
                if voceCodiceNode is None or not voceCodiceNode.text:
                    continue
                # ~hashcode = Voce.find('hashcode').text # il dato c'è, ma per ora non serve
                Scodice = codiceCat + '.' + voceCodiceNode.text
                sottovoci = Voce.find('sottovoci')

                if sottovoci is not None and len(sottovoci):
                    # caso normale: la voce si articola in sottovoci prezzate
                    for Sottovoce in sottovoci:
                        SVcodiceNode = Sottovoce.find('codice')
                        prezzoNode = Sottovoce.find('prezzo')
                        if SVcodiceNode is None or not SVcodiceNode.text or \
                           prezzoNode is None or not prezzoNode.text:
                            # sottovoce priva di codice o prezzo: non utilizzabile
                            continue

                        codice = Scodice + '.' + SVcodiceNode.text
                        try:
                            SVdescNode = Sottovoce.find('descrizione')
                            desc = voce + '\n- ' + SVdescNode.text
                        except:
                            desc = voce

                        umNode = Sottovoce.find('unitaMisura')
                        um = ''
                        if umNode is not None:
                            umCodiceNode = umNode.find('codice')
                            if umCodiceNode is not None and umCodiceNode.text:
                                um = umCodiceNode.text.strip()

                        try:
                            prezzo = float(prezzoNode.text)
                        except ValueError:
                            prezzo = 0.0

                        mdoNode = Sottovoce.find('manodopera')
                        mdo = 0.0
                        if mdoNode is not None and mdoNode.text:
                            try:
                                mdo = float(mdoNode.text) / 100
                            except ValueError:
                                mdo = 0.0
                        if mdo == 0:
                            mdo = ''

                        # un po' di pulizia nel testo
                        # desc = PL.clean_text (desc)

                        # compone l'articolo e lo mette in lista
                        artList[codice] = {
                            'codice': codice,
                            'desc': desc,
                            'um': um,
                            'prezzo': prezzo,
                            'mdo': mdo,
                            'sicurezza': ''
                        }
                else:
                    # la voce non ha sottovoci: è essa stessa l'articolo prezzato
                    # (prezzo/manodopera/unitaMisura direttamente sotto <Voce>)
                    prezzoNode = Voce.find('prezzo')
                    if prezzoNode is None or not prezzoNode.text:
                        # niente prezzo utilizzabile (es. voce puramente
                        # descrittiva/di raggruppamento): si salta
                        continue

                    codice = Scodice
                    desc = voce

                    umNode = Voce.find('unitaMisura')
                    um = ''
                    if umNode is not None:
                        umCodiceNode = umNode.find('codice')
                        if umCodiceNode is not None and umCodiceNode.text:
                            um = umCodiceNode.text.strip()

                    try:
                        prezzo = float(prezzoNode.text)
                    except ValueError:
                        prezzo = 0.0

                    mdoNode = Voce.find('manodopera')
                    mdo = 0.0
                    if mdoNode is not None and mdoNode.text:
                        try:
                            mdo = float(mdoNode.text) / 100
                        except ValueError:
                            mdo = 0.0
                    if mdo == 0:
                        mdo = ''

                    artList[codice] = {
                        'codice': codice,
                        'desc': desc,
                        'um': um,
                        'prezzo': prezzo,
                        'mdo': mdo,
                        'sicurezza': ''
                    }

    # ritorna un dizionario contenente tutto il necessario
    # per costruire l'elenco prezzi
    return {
        'titolo': titolo,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
    }
