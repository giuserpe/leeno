"""
Genera la mappa delle sezioni del manuale FODT.
Usato dalla skill leeno-aggiorna-manuale per mantenere aggiornato
il file MAPPA_SEZIONI.md.

Uso:
    python .agent/skills/leeno-aggiorna-manuale/scripts/genera_mappa.py
"""
import re, os

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..', '..'))
FODT = os.path.join(REPO, 'documentazione', 'MANUALE_LeenO.fodt')
OUT  = os.path.join(os.path.dirname(__file__), '..', 'MAPPA_SEZIONI.md')

with open(FODT, 'r', encoding='utf-8') as f:
    text = f.read()

pattern = r'<text:h[^>]*text:outline-level="(\d+)"[^>]*>(.*?)</text:h>'
headings = re.finditer(pattern, text, re.DOTALL)

lines = []
for m in headings:
    level = int(m.group(1))
    raw = m.group(2)
    # Strip XML tags
    clean = re.sub(r'<[^>]+>', '', raw).strip()
    # Skip entries with newlines (TOC noise) or empty entries
    if not clean or '\n' in clean:
        continue
    pos = m.start()
    line_num = text[:pos].count('\n') + 1
    # Extract bookmark name if present
    bm = re.search(r'text:bookmark-start text:name="([^"]+)"', m.group(0))
    bm_name = bm.group(1) if bm else ''
    indent = '  ' * (level - 1)
    lines.append(f'{indent}- L{line_num} | Lv{level} | {clean} | bookmark: `{bm_name}`')

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('# Mappa delle Sezioni del Manuale LeenO\n\n')
    f.write('Generata automaticamente dal file `documentazione/MANUALE_LeenO.fodt`.\n')
    f.write('Usare questa mappa per individuare rapidamente il punto di inserimento.\n\n')
    f.write('> [!TIP]\n')
    f.write('> Per rigenerare la mappa dopo modifiche al manuale:\n')
    f.write('> `python .agent/skills/leeno-aggiorna-manuale/scripts/genera_mappa.py`\n\n')
    f.write('Formato: `L<riga> | Lv<livello> | <titolo> | bookmark: <nome>`\n\n')
    f.write('\n'.join(lines))
    f.write('\n')

print(f'Mappa generata con {len(lines)} sezioni -> {os.path.abspath(OUT)}')
