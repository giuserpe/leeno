import time
import xml.etree.ElementTree as ET
from io import StringIO
import re

def clean_text(desc):
    desc = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', desc)
    def _remove_illegal_charref(m):
        code = int(m.group(1))
        if code in (9, 10, 13) or code > 31:
            return m.group(0)
        return ''
    desc = re.sub(r'&#(\d+);', _remove_illegal_charref, desc)
    def _remove_illegal_hexref(m):
        code = int(m.group(1), 16)
        if code in (9, 10, 13) or code > 31:
            return m.group(0)
        return ''
    desc = re.sub(r'&#x([0-9a-fA-F]+);', _remove_illegal_hexref, desc)
    return desc

t0 = time.time()
with open(r"W:\_dwg\ULTIMUSFREE\elenchi\Sardegna\2024\SAR24_prezzario_articoli.xml", 'r', encoding='utf-8') as f:
    data = f.read()
print('Read:', time.time()-t0)

t1 = time.time()

artList = {}
superCatList = {}
catList = {}

context = ET.iterparse(StringIO(data), events=("end",))
for event, elem in context:
    tag_name = elem.tag.split('}')[-1]
    if tag_name == 'SAR24_prezzario_articoli':
        def get_text(child_name):
            for child in elem:
                if child.tag.endswith(child_name):
                    return child.text if child.text else ''
            return ''

        codice = get_text('cod')
        if not codice:
            elem.clear()
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
        
        v_clean = ' '.join(sotto_text.upper().split())
        a_clean = ' '.join(desc_text.upper().split())
        if v_clean and a_clean and a_clean.startswith(v_clean):
            desc = clean_text(desc_text)
        else:
            if sotto_text:
                desc = clean_text(sotto_text + '\n' + desc_text)
            else:
                desc = clean_text(desc_text)
                
        um = get_text('um')
        
        prezzo_str = get_text('prezzo_complessivo')
        if not prezzo_str:
            prezzo_str = get_text('prezzo')
        prezzo = float(prezzo_str) if prezzo_str else ''
        
        sic_str = get_text('sicurezza')
        sic = float(sic_str) if sic_str else ''
        
        artList[codice] = {
            'codice': codice,
            'desc': desc,
            'um': um,
            'prezzo': prezzo,
            'mdo': '',
            'sicurezza': sic
        }
        
        elem.clear()

print('Parse:', time.time()-t1)
print('Total elements:', len(artList))
