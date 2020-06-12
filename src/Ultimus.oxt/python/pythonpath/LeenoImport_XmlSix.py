"""
    LeenO - modulo di importazione prezzari formato XML/SIX
"""
from xml.etree.ElementTree import ElementTree

import LeenoUtils
import pyleeno as PL

import Dialogs


def _Fill_Ep(nome, lista_articoli):
    '''
    Crea un nuovo file di computo e lo riempie con la lista articoli
    Ritorna True se OK, False altrimenti
    '''
    progress = None
    #try:
    # creo nuovo file di computo
    PL.New_file.computo(0)

    # e lo compilo
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = nome
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 1).String = nome
    oSheet.getRows().insertByIndex(4, len(lista_articoli))

    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    scarto_colonne = 0  # numero colonne da saltare a partire da sinistra
    scarto_righe = 4  # numero righe da saltare a partire dall'alto
    colonne_lista = len(lista_come_array[1])  # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array)  # numero di righe necessarie per ospitare i dati

    progress = Dialogs.Progress(
        Title="Importazione prezzario XML-SIX",
        Text="Compilazione prezzario in corso",
        MaxVal=righe_lista)
    progress.show()

    riga = 0
    step = 100
    while riga < righe_lista:
        progress.setValue(riga)
        sliced = lista_come_array[riga:riga + step]
        num = len(sliced)
        oRange = oSheet.getCellRangeByPosition(
            scarto_colonne,
            scarto_righe + riga,
            # l'indice parte da 0
            colonne_lista + scarto_colonne - 1,
            scarto_righe + riga + num - 1)
        oRange.setDataArray(sliced)

        riga = riga + step

    oSheet.getRows().removeByIndex(3, 1)
    oDoc.CurrentController.setActiveSheet(oSheet)
    # ~ struttura_Elenco()

    progress.hide()
    Dialogs.Ok(Title="Operazione completata", Text="Importazione eseguita con successo")
    return True

    #except Exception:
    #    if progress is not None:
    #        progress.hideDialog()
    #    Dialogs.Exclamation(Title="Errore", Text="Errore nella compilazione del prezzario\nSegnalare il problema sul sito")
    #    return False


def MENU_Import_Ep_XML_SIX():
    '''
    Routine di importazione di un prezzario XML-SIX in tabella Elenco Prezzi
    del template COMPUTO.
    '''
    newParagraph = "\n"

    filename = Dialogs.FileSelect('Scegli il file XML-SIX da importare', '*.xml')
    if filename is None:
        return

    tree = ElementTree()
    try:
        tree.parse(filename)
    except Exception:
        Dialogs.Exclamation(Title="Errore", Text="Errore nel file XML\nControllarlo e ritentare l'importazione")
        return

    root = tree.getroot()
    prezzario = root.find('{six.xsd}prezzario')
    descrizioni = prezzario.findall('{six.xsd}przDescrizione')
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
        return

    # da qui, se lingua == None importa tutte le lingue presenti
    # altrimenti solo quella specificata

    # estrae il nome
    # se richiesta un lingua specifica, estrae quella
    # altrimenti le estrea tutte e le concatena una dopo l'altra
    nome = ""
    if lingua is None:
        nome = descrizioni[0].attrib['breve']
        for desc in range(1, len(descrizioni)):
            nome = nome + newParagraph + descrizioni[desc].attrib['breve']
    else:
        for desc in descrizioni:
            if desc.attrib['lingua'] == lingua:
                nome = desc.attrib['breve']
                break

    # legge le unità di misura
    # siccome ci interessano soli i simboli e non il resto
    # non serve il processing per le lingue
    units = {}
    umList = prezzario.findall('{six.xsd}unitaDiMisura')
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
    catList = root.findall('{six.xsd}categoriaSOA')
    for cat in catList:
        attr = cat.attrib
        try:
            soaId = attr['soaId']
            soaCategoria = attr['soaCategoria']
            descs = cat.findall('{six.xsd}soaDescrizione')
            text = ""
            for desc in descs:
                descAttr = desc.attrib
                try:
                    descLingua = descAttr['lingua']
                except KeyError:
                    descLingua = None
                if lingua is None or descLingua is None or lingua == descLingua:
                    text = text + descAttr['breve'] + newParagraph
            if text != "":
                text = text[: -len(newParagraph)]

            categorieSOA[soaCategoria] = {'soaId': soaId, 'descrizione': text}
        except KeyError:
            pass

    # infine tiriamo fuori il prezzario
    # utilizziamo le voci 'true' come base per le descrizioni
    # aggiungendo quelle delle voci specializzate '.a, .b...'
    basePrdId = ''
    baseTextBreve = ''
    baseTextEstesa = ''
    products = []
    productList = prezzario.findall('{six.xsd}prodotto')
    for product in productList:
        attr = product.attrib
        try:
            # il codice del prodotto
            prdId = attr['prdId']

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
            valore = ""
            try:
                valAttr = product.find('{six.xsd}prdQuotazione').attrib
                if 'valore' in valAttr:
                    valore = valAttr['valore']
                if valore != "":
                    valore = float(valore)
            except Exception:
                valore = ""

            if valore == 0:
                valore = ""

            # percentuale manodopera
            mdo = ""
            mdoVal = ""
            try:
                mdo = product.find('{six.xsd}incidenzaManodopera').text
                if mdo != "":
                    mdo = float(mdo) / 100
                    if valore != "":
                        mdoVal = valore * mdo
            except Exception:
                mdo = ""
                mdoVal = ""
            if mdo == 0:
                mdo = ""
            if mdoVal == 0:
                mdoVal = ""

            # oneri sicurezza
            sicurezza = ""
            if 'onereSicurezza' in attr:
                sicurezza = attr['onereSicurezza']
            if sicurezza != "":
                sicurezza = float(sicurezza)
            if sicurezza == 0:
                sicurezza = ""

            # per le descrizioni, come sempre... processing a seconda
            # della lingua disponibile / richiesta
            descs = product.findall('{six.xsd}prdDescrizione')
            textBreve = ""
            textEstesa = ""
            for desc in descs:
                descAttr = desc.attrib
                try:
                    descLingua = descAttr['lingua']
                except KeyError:
                    descLingua = None
                if lingua is None or descLingua is None or lingua == descLingua:
                    if 'breve' in descAttr:
                        textBreve = textBreve + descAttr['breve'] + newParagraph
                    if 'estesa' in descAttr:
                        textEstesa = textEstesa + descAttr['estesa'] + newParagraph
            if textBreve != "":
                textBreve = textBreve[: -len(newParagraph)]
            if textEstesa != "":
                textEstesa = textEstesa[: -len(newParagraph)]

            # controlla se la voce è una voce 'base' o una specializzazione
            # della voce base. Il campo 'voce' è totalmente inaffidabile, quindi
            # consideriamo come 'base' delle voci a valore nullo e le azzeriamo
            # ad ogni nuova base e/o cambio di numerazione
            base = (valore == "")
            if not base and not prdId.startswith(basePrdId):
                baseTextBreve = ""
                baseTextEstesa = ""

            # se voce base, tiene buona la descrizione e la salva anche come base
            # per le prossime voci (estese). Per funzionare, questo presuppone che
            # nell' XML le voci estese seguano quella base in ordine, e che tutte le voci
            # siano correttamente etichettate come voce = 'true' se voci base
            # probabilmente si può fare di meglio...
            if base:
                baseTextBreve = textBreve + newParagraph
                baseTextEstesa = textEstesa + newParagraph
                basePrdId = prdId

            if not base and prdId.startswith(basePrdId):
                textBreve = baseTextBreve + textBreve
                textEstesa = baseTextEstesa + textEstesa

            # utilizza solo la descrizione lunga per LeenO
            if len(textBreve) > len(textEstesa):
                textDesc = textBreve
            else:
                textDesc = textEstesa

            # giochino per garantire che la prima stringa abbia una lunghezza minima
            # in modo che LO formatti correttamente la cella
            minLen = 130
            splitted = textDesc.split(newParagraph)
            if len(splitted) > 1:
                spl0 = splitted[0]
                spl1 = splitted[1]
                if len(spl0) + len(spl1) < minLen:
                    dl = minLen - len(spl0) - len(spl1)
                    spl0 = spl0 + dl * " "
                    textDesc = spl0 + newParagraph + spl1
                    for txt in splitted[2:]:
                        textDesc = textDesc + newParagraph + txt

            # inserisce nella lista
            # per risparmiare tempo la lista è creata direttamente come array
            # corrispondente ai dati da inserire nel foglio dell' elenco prezzi
            products.append([prdId, textDesc, um, sicurezza, valore, mdo, mdoVal])
        except KeyError:
            pass

    # compila la tabella
    _Fill_Ep(nome, products)
    PL.autoexec()
