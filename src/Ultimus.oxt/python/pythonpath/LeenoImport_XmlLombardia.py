import re
import LeenoImport
import xml.etree.ElementTree as ET
import LeenoDialogs as DLG
import pyleeno as PL

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
    data = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', data)

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    root = LeenoImport.stripXMLNamespaces(data)

    voci = root.find('voci/voci')
    rifvoce = voci.find('riferimenti_voce')

    try:
        titolo = f"{rifvoce.find('autore').text} {rifvoce.find('invigore').text} {rifvoce.find('anno').text}"
    except AttributeError:
        titolo = f"{rifvoce.find('autore').text} {rifvoce.find('anno').text}"
    
    try:
        titolo += f"_{rifvoce.find('edizione').text}"
    except AttributeError:
        pass
    
    artList = {}
    superCatList = {}
    catList = {}

    # prendo il suffisso da aggiungere al codice categorie
    try:
        suff = root[0][0].getchildren()[1].attrib['CMPcodifica_voce'][:8]
    except:
        suff = root[0][0].getchildren()[1].attrib['codice_voce'][:8]

    voci = root.find('voci')
    for voce in voci:
        dettaglio_voce = list(voce)[1]
        ###
        # estrae supercategoria e categoria
        try:
            codiceSuperCat = suff + dettaglio_voce.attrib['codifica_I_livello_voce']
            superCat = dettaglio_voce.attrib['declaratoria_I_livello_voce']
            if not codiceSuperCat in superCatList:
                superCatList[codiceSuperCat] = superCat

            codiceSuperCat = suff + dettaglio_voce.attrib['codifica_II_livello_voce']
            superCat = dettaglio_voce.attrib['declaratoria_II_livello_voce']
            if not codiceSuperCat in superCatList:
                superCatList[codiceSuperCat] = superCat

            codiceSuperCat = suff + dettaglio_voce.attrib['codifica_III_livello_voce']
            superCat = dettaglio_voce.attrib['declaratoria_III_livello_voce']
            if not codiceSuperCat in superCatList:
                superCatList[codiceSuperCat] = superCat

            codiceSuperCat = suff + dettaglio_voce.attrib['codifica_IV_livello_voce']
            superCat = dettaglio_voce.attrib['declaratoria_IV_livello_voce']
            if not codiceSuperCat in superCatList:
                superCatList[codiceSuperCat] = superCat

            codiceSuperCat = suff + dettaglio_voce.attrib['codifica_V_livello_voce']
            superCat = dettaglio_voce.attrib['declaratoria_V_livello_voce']

            if not codiceSuperCat in superCatList:
                superCatList[codiceSuperCat] = superCat
                codiceSuperCat = suff + dettaglio_voce.attrib['codifica_VI_livello_voce']
                superCat = dettaglio_voce.attrib['declaratoria_VI_livello_voce']
            if not codiceSuperCat in superCatList:
                superCatList[codiceSuperCat] = superCat
        except:
            pass
        ###
        try:
            codice = dettaglio_voce.attrib['CMPcodifica_voce']
        except KeyError:
            codice = dettaglio_voce.attrib['codice_voce']
        
        desc = dettaglio_voce.find('declaratoria_voce').text
        
        try:
            um = dettaglio_voce.attrib['udm_voce']
        except KeyError:
            um = dettaglio_voce.attrib['unita_misura_voce']
        
        prezzo = float(dettaglio_voce.attrib['prezzo_voce'])
        
        risorse = dettaglio_voce.find('risorse')
        mdo = ''
        if risorse is not None:
            for el in risorse:
                if el.attrib['tipologia_risorsa'] == "MANODOPERA":
                    mdo = float(el.attrib['perc_importo_tipo_risorsa'])
                    break
        
        try:
            mdo = float(dettaglio_voce.attrib['rapporto_RU_voce']) / 100
        except (KeyError, ValueError):
            pass

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
        'articoli': artList
    }

########################################################################

def parseXML1(data, defaultTitle=None):
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
    titolo = root.items()[0][-1].split('.')[0]

    # ~ voci = root.find('Parte3')
    voci = list(root.getchildren())
            # ~ 'codice': codice,
            # ~ 'desc': desc,
            # ~ 'um': um,
            # ~ 'prezzo': prezzo,
            # ~ 'mdo': mdo,
            # ~ 'sicurezza': ''
    artList = {}
    superCatList = {}
    catList = {}
    madre = ''
    for voce in voci:
        codice = voce.find('Codice').text
        if ' - ' in codice:
            codice = codice.split(' - ')[0]
            # ~ DLG.chi(codice)
            DLG.chi(voce.find('Codice').text[len(codice):])
            desc = voce.find('Codice').text[len(codice):]
            # ~ return
        try:
            desc = PL.clean_text(voce.find('Declaratoria').text)
        except Exception as e:
            # ~ DLG.chi(f'Errore: {e} code: {codice}')
            desc = ''
        try:
            um = voce.find('UM').text
            desc = madre + '\n' + desc
        except:
            um = ''
            madre = desc
        try:
            prezzo = float(voce.find('Prezzo').text)
        except:
            prezzo = ''
        try:
            mdo = float(voce.find('Rapporto_RU').text)
        except:
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
        'articoli': artList
    }
