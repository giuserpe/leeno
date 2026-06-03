"""
    LeenO - modulo parser XML per il formato XML-SIX
"""
import LeenoImport_XmlToscana
import LeenoImport
import pyleeno as PL

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

    # Se il file contiene il formato SAR24 (dataroot / SAR24_prezzario_articoli)
    if '<SAR24_prezzario_articoli>' in data:
        import xml.etree.ElementTree as ET
        from io import StringIO
        
        artList = {}
        superCatList = {}
        catList = {}
        
        context = ET.iterparse(StringIO(data), events=("start", "end"))
        context = iter(context)
        # Ottiene il nodo radice (dataroot) al primo evento
        _, root = next(context)

        for event, elem in context:
            tag_name = elem.tag.split('}')[-1]
            if event == 'end' and tag_name == 'SAR24_prezzario_articoli':
                def get_text(child_name):
                    for child in elem:
                        if child.tag.endswith(child_name):
                            return child.text if child.text else ''
                    return ''

                codice = get_text('cod')
                if not codice:
                    elem.clear()
                    root.clear()
                    continue
                
                codiceSplit = codice.split('.')
                if len(codiceSplit) >= 2:
                    codiceSuperCat = codiceSplit[0]
                    codiceCat = codiceSuperCat + '.' + codiceSplit[1]
                else:
                    codiceSuperCat = codice
                    codiceCat = codice
                    
                if codiceSuperCat not in superCatList:
                    superCatList[codiceSuperCat] = f"{get_text('famiglia')} - {get_text('capitolo')}".strip(" -")
                    
                if codiceCat not in catList:
                    catList[codiceCat] = get_text('sottocapitolo')
                    
                desc_text = get_text('descrizione')
                sotto_text = get_text('sottocapitolo')
                
                # Evita duplicazioni tra sottocapitolo e descrizione
                v_clean = ' '.join(sotto_text.upper().split())
                a_clean = ' '.join(desc_text.upper().split())
                if v_clean and a_clean and a_clean.startswith(v_clean):
                    desc = PL.clean_text(desc_text)
                else:
                    if sotto_text:
                        desc = PL.clean_text(sotto_text + '\n' + desc_text)
                    else:
                        desc = PL.clean_text(desc_text)
                        
                um = get_text('um')
                
                # Gestione robusta dei prezzi per evitare crash silenziosi in LibreOffice
                try:
                    prezzo_str = get_text('prezzo_complessivo')
                    if not prezzo_str:
                        prezzo_str = get_text('prezzo')
                    prezzo_str = prezzo_str.replace(',', '.')
                    prezzo = float(prezzo_str) if prezzo_str.strip() else ''
                except ValueError:
                    prezzo = ''
                
                try:
                    sic_str = get_text('sicurezza').replace(',', '.')
                    sic = float(sic_str) if sic_str.strip() else ''
                except ValueError:
                    sic = ''
                
                artList[codice] = {
                    'codice': codice,
                    'desc': desc,
                    'um': um,
                    'prezzo': prezzo,
                    'mdo': '',
                    'sicurezza': sic
                }
                
                # Libera memoria svuotando completamente il nodo appena processato dal root
                elem.clear()
                root.clear()
            
        return {
            'titolo': "Elenco prezzi - Regione Sardegna - anno 2024",
            'superCategorie': superCatList,
            'categorie': catList,
            'articoli': artList
        }

    # il formato (ad es. Sardegna 2025 EASY) è quello della toscana
    return LeenoImport_XmlToscana.parseXML(data, defaultTitle)

