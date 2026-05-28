#!/usr/bin/env python3
"""Estrae data, dimensione e nome dei file .oxt da webdav.xml (PROPFIND)."""
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime

NS = {
    'd': 'DAV:',
    'oc': 'http://owncloud.org/ns',
    'nc': 'http://nextcloud.org/ns',
}

try:
    tree = ET.parse('webdav.xml')
    root = tree.getroot()
except Exception as e:
    print(f"Errore nel parsing di webdav.xml: {e}")
    # Crea un file vuoto per evitare crash successivi
    with open('oxt_list.txt', 'w') as f:
        pass
    import sys
    sys.exit(0)

entries = []
for resp in root.findall('d:response', NS):
    href = resp.findtext('d:href', '', NS)
    if not href or not href.lower().endswith('.oxt'):
        continue
    
    name = href.rstrip('/').split('/')[-1]
    
    # Cerca globalmente all'interno di response per evitare problemi con propstat multipli
    size_b = resp.findtext('.//d:getcontentlength', default='0', namespaces=NS)
    if not size_b: size_b = '0'
    
    date_raw = resp.findtext('.//d:getlastmodified', default='', namespaces=NS)
    if not date_raw: date_raw = ''
    
    try:
        dt = parsedate_to_datetime(date_raw)
        date = dt.strftime('%Y-%m-%d %H:%M')
    except Exception:
        date = date_raw[:10] if date_raw else 'N/A'
        
    try:
        size_mb = f"{int(size_b) / 1048576:.1f}MB"
    except Exception:
        size_mb = 'N/A'
        
    entries.append((date, size_mb, name))

entries.sort(reverse=True)

with open('oxt_list.txt', 'w') as f:
    for date, size, name in entries[:10]:
        f.write(f"{date} {size} {name}\n")

print(f"Scritti {min(len(entries), 10)} file in oxt_list.txt")
