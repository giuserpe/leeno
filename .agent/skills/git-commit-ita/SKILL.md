---
name: git-commit-ita
description: >
  Genera messaggi di commit Git seguendo lo standard Conventional Commits
  con descrizione in italiano. Usare quando si vuole fare un commit,
  scrivere o formattare un messaggio di commit.
---

# Git Commit – Conventional Commits in Italiano

## Formato
```
<tipo>(<scope opzionale>): <descrizione in italiano>

[corpo opzionale]
```

## Tipi

| Tipo | Quando |
|------|--------|
| `feat` | Nuova funzionalità |
| `fix` | Correzione bug |
| `docs` | Solo documentazione |
| `style` | Formattazione, no logica |
| `refactor` | Refactoring senza feat né fix |
| `perf` | Miglioramento prestazioni |
| `test` | Aggiunta/modifica test |
| `chore` | Manutenzione, dipendenze |
| `revert` | Annullamento commit |

## Regole

1. Descrizione in **italiano**, imperativo presente
2. Max 72 caratteri nell'intestazione
3. Nessun punto finale
4. Breaking change: aggiungi `!` dopo il tipo

## Esempi
```
feat(calc): aggiunge supporto macro per il foglio imposte
fix(ui): corregge allineamento colonne nella tabella riassuntiva
docs: aggiorna README con istruzioni installazione estensione
chore: aggiorna dipendenze Python
```

## Procedura

1. Esegui `git diff --cached` per analizzare le modifiche staged
2. Determina il tipo più appropriato
3. Proponi il comando completo:
   `git commit -m "<tipo>(<scope>): <descrizione>"`
