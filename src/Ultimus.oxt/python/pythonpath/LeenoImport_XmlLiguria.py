import xml.etree.ElementTree as ET
import LeenoImport

# Importiamo l'oggetto di logging e le utility di LeenO (es. PL se è il modulo dei prezzari)
# In base al tuo LeenoImport, usiamo le funzioni di pulizia testo native se necessario.

def parseXML(data, defaultTitle=None):
    '''
    Estrae dal file XML i dati dell'elenco prezzi Liguria.
    Il formato restituito è compatibile con le routine di scrittura di LeenO.
    '''
    # Pulizia del testo tramite la funzione passata o interna
    # Se PL non è importato, si può usare una regex o un helper di LeenO
    if 'PL' in globals():
        data = PL.clean_text(data)

    # Elimina i namespaces dai dati ed ottiene l'elemento radice
    root = LeenoImport.stripXMLNamespaces(data)

    intestazione = root.find('intestazione')
    autore = intestazione.attrib['autore']

    dettaglio = intestazione.find('dettaglio')
    anno = dettaglio.attrib['anno']
    area = dettaglio.attrib['area']

    copyright = intestazione.find('copyright')
    ccType = copyright.attrib['tipo']
    ccDesc = copyright.attrib['descrizione']

    # Compone il titolo dell'Elenco Prezzi
    titolo = f"Elenco prezzi - {area} - anno {anno}\nCopyright: {ccType} - {ccDesc}"

    contenuto = root.find('Contenuto')
    articoli = contenuto.findall('Articolo')

    artList = {}
    superCatList = {}
    catList = {}

    for articolo in articoli:
        codice = articolo.attrib['codice']

        # Divide il codice per ottenere i codici di supercategoria e categoria
        codiceSplit = codice.split('.')
        codiceSuperCat = codiceSplit[0]
        
        # Gestione di sicurezza nel caso il codice non abbia abbastanza sotto-livelli
        if len(codiceSplit) > 1:
            codiceCat = codiceSuperCat + '.' + codiceSplit[1]
        else:
            codiceCat = codiceSuperCat

        # Estrae supercategoria e categoria
        superCat = articolo.find('tipo').text if articolo.find('tipo') is not None else ''
        cat = articolo.find('capitolo').text if articolo.find('capitolo') is not None else ''

        # Inserimento nelle liste uniche
        if codiceSuperCat not in superCatList:
            superCatList[codiceSuperCat] = superCat
        if codiceCat not in catList:
            catList[codiceCat] = cat

        voce = articolo.find('voce').text or ''
        art_text = articolo.find('articolo').text or ''
        desc = f"{voce}\n- {art_text}".strip()

        # Pulizia Unità di Misura (rimuove eventuali parentesi)
        um_node = articolo.find('um')
        um = um_node.text.split('(')[-1][:-1] if um_node is not None and um_node.text else ''

        # Gestione e pulizia del prezzo (rimozione separatori migliaia orfani)
        prezzo_node = articolo.find('prezzo')
        prezzo = 0.0
        if prezzo_node is not None:
            prezzo_val = prezzo_node.attrib.get('valore', '0')
            if '.' not in prezzo_val:
                prezzo_val = prezzo_val + '.0'
            prSplit = prezzo_val.split('.')
            prezzo_pulito = ''.join(prSplit[0:-1]) + '.' + prSplit[-1]
            try:
                prezzo = float(prezzo_pulito)
            except ValueError:
                prezzo = 0.0

        # Manodopera e Sicurezza
        mdo = float(articolo.find('mo').text or 0) / 100
        oneriSic = float(articolo.find('sicurezza').text or 0)
        if oneriSic == 0:
            oneriSic = ''

        # Compone l'articolo e lo inserisce nel dizionario
        artList[codice] = {
            'codice': codice,
            'desc': desc,
            'um': um,
            'prezzo': prezzo,
            'mdo': mdo,
            'sicurezza': oneriSic
        }

    return {
        'titolo': titolo,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
    }


def importa_liguria(PrezzarioInstance, xml_path):
    '''
    Funzione principale di interfaccia per LeenO.
    Riceve l'istanza del prezzario corrente e il percorso del file XML ligure.
    '''
    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            data = f.read()
    except UnicodeDecodeError:
        with open(xml_path, 'r', encoding='latin-1') as f:
            data = f.read()

    # Eseguiamo il parsing specifico per la Liguria
    dati = parseXML(data)
    
    # Assegniamo il titolo formattato all'istanza del prezzario di LeenO
    PrezzarioInstance.titolo = dati['titolo']
    
    # Ordiniamo i codici degli articoli per garantire la sequenzialità logica nell'importazione
    codici_ordinati = sorted(dati['articoli'].keys())
    
    supercat_inserite = set()
    cat_inserite = set()
    
    # Utilizziamo i metodi nativi dell'istanza Prezzario di LeenO per scrivere nel foglio Calc
    for codice in codici_ordinati:
        codiceSplit = codice.split('.')
        codiceSuperCat = codiceSplit[0]
        codiceCat = codiceSuperCat + '.' + codiceSplit[1] if len(codiceSplit) > 1 else codiceSuperCat
        
        # 1. Inserimento Supercategoria (Capitolo principale)
        if codiceSuperCat not in supercat_inserite:
            if codiceSuperCat in dati['superCategorie']:
                desc_super = dati['superCategorie'][codiceSuperCat]
                # Se LeenO ha una funzione nativa per i capitoli/voci macro, usala qui
                PrezzarioInstance.aggiungi_voce_macro(codiceSuperCat, desc_super.upper()) 
                supercat_inserite.add(codiceSuperCat)
                
        # 2. Inserimento Categoria (Sub-capitolo)
        if codiceCat not in cat_inserite:
            if codiceCat in dati['categorie']:
                desc_cat = dati['categorie'][codiceCat]
                PrezzarioInstance.aggiungi_voce_macro(codiceCat, desc_cat)
                cat_inserite.add(codiceCat)
                
        # 3. Inserimento Articolo di dettaglio
        art = dati['articoli'][codice]
        
        # Utilizziamo il metodo standard di LeenO per iniettare la riga articolo compilata
        # Adattando i campi al dizionario atteso dal vostro costruttore di righe
        PrezzarioInstance.scrivi_articolo(
            codice=art['codice'],
            descrizione=art['desc'],
            um=art['um'],
            prezzo=art['prezzo'],
            mdo=art['mdo'],
            sicurezza=art['sicurezza']
        )