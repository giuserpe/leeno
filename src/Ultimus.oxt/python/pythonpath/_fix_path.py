f = 'pyleeno.py'
lines = open(f, encoding='utf-8').read().splitlines()

# Sostituisci righe 12690-12711 (0-based: 12689-12710) con la nuova funzione
new_func = [
    "def MENU_importa_da_dcf():",
    "    try:",
    "        path = Dialogs.FileSelect('Seleziona file DCF...', '*.dcf', 0)",
    "        if not path:",
    "            return",
    "",
    "        doc = parse_dcf(path)",
    "        rows = computo_con_descrizioni(doc)",
    "",
    "        totale_globale = sum(r['importo'] for r in rows)",
    "",
    "        from collections import defaultdict",
    "        per_cat = defaultdict(list)",
    "        for r in rows:",
    "            chiave = (r['super_cat'] or '\u2014', r['categoria'] or '\u2014')",
    "            per_cat[chiave].append(r)",
    "",
    "        msg = f'File: {path}\\n'",
    "        msg += f'Voci totali: {len(rows)}\\n'",
    "        msg += f'Totale complessivo: euro {totale_globale:,.2f}\\n'",
    "        msg += '=' * 60 + '\\n\\n'",
    "",
    "        for (spcat, cat), voci in sorted(per_cat.items()):",
    "            subtot = sum(v['importo'] for v in voci)",
    "            msg += f'[{spcat}] > {cat}  ({len(voci)} voci, euro {subtot:,.2f})\\n'",
    "            for r in voci:",
    "                msg += (f\"  {r['tariffa']:18s} {r['um']:4s}\"",
    "                        f\" Qt={r['quantita']:8.3f}  euro {r['importo']:>10,.2f}\"",
    "                        f\"  {r['descrizione'][:45]}\\n\")",
    "            msg += '\\n'",
    "",
    "        DLG.chi(msg)",
    "    except Exception as e:",
    "        DLG.chi(f'ERRORE in MENU_importa_da_dcf:\\n{e}')",
]

# Sostituisci le righe 12689-12710 (indici 0-based)
start_idx = 12689   # 0-based indice della riga 12690
end_idx   = 12711   # 0-based indice della riga 12712 (esclusa, era la riga vuota prima di MENU_debug)

lines_new = lines[:start_idx] + new_func + lines[end_idx:]
print(f"Righe originali: {len(lines)}, Righe nuove: {len(lines_new)}")

open(f, 'w', encoding='utf-8', newline='\r\n').write('\n'.join(lines_new))
print("OK: file scritto")
