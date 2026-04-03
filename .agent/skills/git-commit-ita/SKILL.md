---
name: git-commit-ita
description: >
  Genera messaggi di commit Git seguendo lo standard Conventional Commits
  con descrizione in italiano. Ottimizzato per il progetto LeenO con scope
  specifici e procedura di analisi dei file staged.
---

# Git Commit – Conventional Commits in Italiano (LeenO)

## Formato
```
<tipo>(<scope>): <descrizione in italiano>

[corpo opzionale: spiega il PERCHÉ, non il COSA]
```

## Tipi

| Tipo | Quando |
|------|--------|
| `feat` | Nuova funzionalità |
| `fix` | Correzione bug |
| `docs` | Solo documentazione |
| `style` | Formattazione, spazi, punti e virgola mancanti (no logica) |
| `refactor` | Modifica del codice che non corregge bug né aggiunge funzionalità |
| `perf` | Miglioramento prestazioni |
| `test` | Aggiunta/modifica test |
| `chore` | Manutenzione, aggiornamento dipendenze, versioning, build |
| `revert` | Annullamento di un commit precedente |

## Scope Suggeriti (LeenO)

Identifica l'area principale colpita dalle modifiche:

- `core`: Logica principale (`pyleeno.py`, `LeenoGlobals.py`, ecc.)
- `ui`: Interfaccia utente (`.xhp`, `.xlb`, dialoghi in Python)
- `contab`, `computo`, `variante`, `giornale`: Modulo specifico in `pythonpath`
- `import`: Filtri di importazione (`LeenoImport_*.py`)
- `icons`: Icone e risorse grafiche (`icons/`, SVG/PNG)
- `meta`: Metadati estensione (`description.xml`, `.xcu`)
- `template`: Modifiche ai modelli di documento
- `docs`: Manuale PDF o documentazione tecnica

## Regole d'Oro

1. **Lingua**: Descrizione in **italiano**, imperativo presente (es. "aggiunge", non "aggiunto")
2. **Lunghezza**: Max 72 caratteri per l'intestazione
3. **Punteggiatura**: Nessun punto finale nell'intestazione
4. **Breaking Change**: Aggiungi `!` dopo il tipo (es. `feat!: ...`) e descrivi in `BREAKING CHANGE:` nel corpo
5. **Separazione**: Se le modifiche riguardano aree troppo diverse, suggerisci commit separati

## Procedura Operativa

1. **Analisi Stato**: Esegui `git status` per vedere quali file sono staged e quali no.
2. **Analisi Modifiche**: Esegui `git diff --cached` per esaminare nel dettaglio il codice modificato.
3. **Identificazione Scope**: Scegli lo scope più calzante in base ai file modificati.
4. **Draft Messaggio**: Componi l'intestazione. Se la modifica non è auto-esplicativa, aggiungi un paragrafo di corpo dopo una riga vuota.
5. **Proponi Comando**: Mostra il comando finale: `git commit -m "..."` o `git commit -e` se serve un corpo esteso.

## Esempi

- `feat(computo): aggiunge calcolo automatico oneri sicurezza`
- `fix(ui): corregge refresh tabella dopo inserimento voce`
- `refactor(import): ottimizza parsing file XPWE`
- `chore(meta): bump versione a 3.25.x`
- `docs: aggiorna istruzioni nel manuale per il nuovo listino`
