# Lezioni apprese – Icone

Regole operative estratte da `AGENTS.md`, che conserva le tre invarianti (stessa geometria in `svg/` e `scuro/`, modifica in modalità binaria, `.bmp` come testo). Il testo qui sotto è quello originale. La specifica del design system è in `ICONS_DESIGN_SYSTEM.md`.

- La specifica completa del design system (filosofia, primitive geometriche, palette colori, regole di export) vive in `documentazione/ICONS_DESIGN_SYSTEM.md`. Consultarla prima di creare o modificare icone: contiene, tra l'altro, la sezione 15 "Canvas di Export 48×48 px con Padding Azzerato", che descrive il crop del `viewBox` calcolato per-icona in uso dalla generazione corrente.
- Il disegno segue la griglia master 24×24 con margine di sicurezza 2px; l'export finale usa invece un canvas quadrato 48×48 con `viewBox` ritagliato individualmente sul bounding box del contenuto (nessun crop fisso uguale per tutte le icone).
- Per icone con badge d'angolo o elementi vicini al bordo, preferire una revisione icona per icona invece di un'operazione bulk automatica: il rischio di danno visivo (badge tagliato, asimmetria) è più alto che nelle icone semplici.
- **Per il calcolo del bounding box in fase di export, usare la pipeline `rsvg-convert` + PIL**, che renderizza correttamente anche gli elementi di testo. Evitare `svgelements` per questo calcolo: restituisce bounding box di dimensione zero per i nodi di testo, producendo crop errati.
