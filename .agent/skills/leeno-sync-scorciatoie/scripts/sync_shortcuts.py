import xml.etree.ElementTree as ET
import zipfile
import os
import shutil

# This script is located in .agent/skills/leeno-sync-scorciatoie/scripts/
# The OXT source is at ../../../../src/Ultimus.oxt/ relative to this script's directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "..", "..", ".."))
BASE_DIR = os.path.join(REPO_ROOT, "src", "Ultimus.oxt")

ACCEL_FILE = os.path.join(BASE_DIR, "Accelerators.xcu")
ADDONS_FILE = os.path.join(BASE_DIR, "Addons.xcu")
TEMPLATE_FILE = os.path.join(BASE_DIR, "template", "leeno", "Computo_LeenO.ods")

# Temporary working directory inside the repo (standard practice for LeenO tools)
TEMP_DIR = os.path.join(REPO_ROOT, "temp_sync_shortcuts")

# Namespaces
NS = {
    'oor': 'http://openoffice.org/2001/registry',
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'calcext': 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
}

for prefix, uri in NS.items():
    if prefix != '':
        ET.register_namespace(prefix, uri)

def format_shortcut_name(raw):
    parts = raw.split('_')
    key = parts[0]
    mods = parts[1:]
    key_map = {'DELETE': 'Del', 'INSERT': 'Ins', 'COMMA': ',', 'ADD': '+', 'SUBTRACT': '-'}
    res = []
    if 'MOD1' in mods: res.append('Ctrl')
    if 'SHIFT' in mods: res.append('Shift')
    if 'MOD2' in mods: res.append('Alt')
    res.append(key_map.get(key, key).capitalize())
    return '+'.join(res)

def sync():
    if not os.path.exists(TEMP_DIR):
        os.makedirs(TEMP_DIR)
    
    # 1. Parse Accelerators
    accel_shortcuts = []
    accel_tree = ET.parse(ACCEL_FILE)
    module_node = accel_tree.getroot().find(".//node[@oor:name='com.sun.star.sheet.SpreadsheetDocument']", NS)
    if module_node is not None:
        for node in module_node.findall("node", NS):
            name = node.get(f'{{{NS["oor"]}}}name')
            val = node.find(".//prop[@oor:name='Command']/value", NS)
            if val is not None:
                accel_shortcuts.append({'name': name, 'cmd': val.text})

    # 2. Parse Addons Labels
    labels = {}
    addons_tree = ET.parse(ADDONS_FILE)
    for node in addons_tree.getroot().findall(".//node"):
        u = node.find("prop[@oor:name='URL']/value", NS)
        t = node.find("prop[@oor:name='Title']/value", NS)
        if u is not None and t is not None:
            labels[u.text] = t.text

    # 3. Unzip ODS
    unzipped_path = os.path.join(TEMP_DIR, "ods_content")
    if os.path.exists(unzipped_path): shutil.rmtree(unzipped_path)
    with zipfile.ZipFile(TEMPLATE_FILE, 'r') as zip_ref:
        zip_ref.extractall(unzipped_path)
    
    content_xml = os.path.join(unzipped_path, "content.xml")
    tree = ET.parse(content_xml)
    root = tree.getroot()
    sheet = root.find(".//table:table[@table:name='Scorciatoie']", NS)
    
    if sheet is None:
        print("Scorciatoie sheet not found!")
        return

    # Categorize
    groups = {'CTRL+SHIFT': [], 'CTRL': [], 'SHIFT': [], 'ALT': [], 'OTHER': []}
    for s in accel_shortcuts:
        name, cmd = s['name'], s['cmd']
        readable = format_shortcut_name(name)
        desc = labels.get(cmd, cmd.split('.')[-1].replace('MENU_', '').replace('_', ' ').capitalize())
        macro = cmd.replace("service:org.giuseppe-vizziello.leeno.dispatcher?", "").replace("vnd.sun.star.script:UltimusFree2.", "").split('?')[0]
        
        target = 'OTHER'
        if 'MOD1' in name and 'SHIFT' in name: target = 'CTRL+SHIFT'
        elif 'MOD1' in name: target = 'CTRL'
        elif 'SHIFT' in name: target = 'SHIFT'
        elif 'MOD2' in name: target = 'ALT'
        groups[target].append((readable, desc, macro))

    # Build rows
    rows = sheet.findall('table:table-row', NS)[:6] # Keep header
    
    def add_sec(title, items):
        if not items: return
        # Section title
        r = ET.Element(f'{{{NS["table"]}}}table-row', {f'{{{NS["table"]}}}style-name': 'ro33'})
        r.append(ET.Element(f'{{{NS["table"]}}}table-cell'))
        c = ET.SubElement(r, f'{{{NS["table"]}}}table-cell', {f'{{{NS["table"]}}}style-name': 'ce594', f'{{{NS["office"]}}}value-type': 'string', f'{{{NS["table"]}}}number-columns-spanned': '3'})
        ET.SubElement(c, f'{{{NS["text"]}}}p').text = title
        r.append(ET.Element(f'{{{NS["table"]}}}covered-table-cell', {f'{{{NS["table"]}}}number-columns-repeated': '2'}))
        r.append(ET.Element(f'{{{NS["table"]}}}table-cell', {f'{{{NS["table"]}}}number-columns-repeated': '16380'}))
        rows.append(r)
        
        for short, d, m in sorted(items):
            row = ET.Element(f'{{{NS["table"]}}}table-row')
            row.append(ET.Element(f'{{{NS["table"]}}}table-cell'))
            c1 = ET.SubElement(row, f'{{{NS["table"]}}}table-cell', {f'{{{NS["office"]}}}value-type': 'string', f'{{{NS["table"]}}}style-name': 'ce594'})
            p1 = ET.SubElement(c1, f'{{{NS["text"]}}}p')
            ET.SubElement(p1, f'{{{NS["text"]}}}span', {f'{{{NS["text"]}}}style-name': 'T4'}).text = short
            
            c2 = ET.SubElement(row, f'{{{NS["table"]}}}table-cell', {f'{{{NS["office"]}}}value-type': 'string'})
            ET.SubElement(c2, f'{{{NS["text"]}}}p').text = d
            
            c3 = ET.SubElement(row, f'{{{NS["table"]}}}table-cell', {f'{{{NS["office"]}}}value-type': 'string'})
            p3 = ET.SubElement(c3, f'{{{NS["text"]}}}p')
            ET.SubElement(p3, f'{{{NS["text"]}}}span', {f'{{{NS["text"]}}}style-name': 'T5'}).text = m
            row.append(ET.Element(f'{{{NS["table"]}}}table-cell', {f'{{{NS["table"]}}}number-columns-repeated': '16380'}))
            rows.append(row)
        # Empty row
        rows.append(ET.Element(f'{{{NS["table"]}}}table-row'))

    add_sec("Combinazioni con CTRL", groups['CTRL'])
    add_sec("Combinazioni con SHIFT", groups['SHIFT'])
    add_sec("Combinazioni con CTRL+SHIFT", groups['CTRL+SHIFT'])
    add_sec("Combinazioni con ALT", groups['ALT'])

    for r in list(sheet.findall('table:table-row', NS)): sheet.remove(r)
    for r in rows: sheet.append(r)
    
    tree.write(content_xml, encoding='utf-8', xml_declaration=True)

    # Re-zip
    bak_file = TEMPLATE_FILE + ".bak"
    shutil.copy2(TEMPLATE_FILE, bak_file)
    with zipfile.ZipFile(TEMPLATE_FILE, 'w', zipfile.ZIP_DEFLATED) as zip_out:
        for root, dirs, files in os.walk(unzipped_path):
            for file in files:
                fp = os.path.join(root, file)
                zip_out.write(fp, os.path.relpath(fp, unzipped_path))
    
    shutil.rmtree(TEMP_DIR)
    print(f"Aggiornamento completato. Backup creato in {os.path.basename(bak_file)}")

if __name__ == "__main__":
    sync()
