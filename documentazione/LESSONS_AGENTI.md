# Lezioni apprese: lavoro delegato agli agenti e revisione delle loro PR

Esperienza di ottobre-novembre 2026 con Jules (Google) su pulizia del codice, test e docstring di `pythonpath/`: oltre 25 PR, sei tipi di task. Il flusso vale per qualsiasi agente. Jules è in dismissione: secondo fonti di stampa le sessioni attive si possono completare fino all'11 novembre 2026 e il servizio si spegne il 31 dicembre 2026 (da riverificare).

## Come assegnare un task

- Un task per modulo (o per pochi moduli simili), una PR, un commit per modulo. Una PR su decine di moduli non si rivede.
- Il prompt deve poter essere incollato così com'è: nessun segnaposto da riempire a mano, oppure l'elenco di cosa sostituire in testa.
- Il prompt dice cosa fare, cosa NON fare, quali controlli eseguire, e chiede solo numeri ottenuti dai comandi.
- Ordine per rischio crescente: CI, test, rimozione di import e variabili inutilizzate, docstring. I refactor di funzioni da centinaia di righe e tutto ciò che dipende dal comportamento di Calc non si delegano: l'agente non può provarli in LibreOffice.
- Le liste di righe da modificare si filtrano con l'AST prima di darle all'agente: 33 istruzioni isolate stimate erano 10, le altre 22 stavano dentro un `try` (sonde volute).

## Errori ricorrenti degli agenti

| Errore | Esempio reale | Rimedio |
| --- | --- | --- |
| PR con commit vuoti e descrizione dettagliata | #93: due commit senza modifiche, numeri inventati | `git diff --stat <base> <head>` prima di leggere la descrizione |
| Note o classificazioni attribuite alla funzione sbagliata | audit degli except: righe e funzioni corrette, colonna Note sfasata in `pyleeno.py` | con l'AST, ogni nota deve citare un nome presente nel corpo del `try` |
| Docstring di un wrapper dedotta dal nome | `LeenoBasicBridge`: 4 su 15 imprecise (`autoexec`, `autoexec_off`, `sproteggi_sheet_TUTTE`, `Rinumera_TUTTI_Capitoli2`) | leggere la funzione richiamata, anche in un altro modulo |
| Istruzione isolata lasciata dopo la rimozione di un'assegnazione | `getCellRangeByPosition(...)`, `set(...)`, `getComponentContext()` | togliere la riga se è una lettura; tenere la chiamata solo se ha effetti |
| Import rimasto orfano | `import pyleeno as PL` in `Calendario.py` | rieseguire pyflakes dopo ogni PR di pulizia |
| Scostamenti dalla regola "solo le righe indicate" | `except:` nudo diventato `except Exception`, commento tolto | confrontare il diff riga per riga, non solo l'AST |

## Checklist di revisione di una PR di agente

1. `git fetch origin pull/<N>/head:pr<N>` e `git merge-base origin/dev pr<N>`: se coincide con la testa della PR, è già su `dev`.
2. `python3 scripts/verifica_diff.py <base> <head>`: AST, line-ending, nomi non definiti, except silenziosi, nomi rimossi ancora usati.
3. I nomi non definiti di pyflakes restano invariati; nelle PR di pulizia il totale degli avvisi scende.
4. Per ogni nome di modulo rimosso: cercarlo come `Modulo.nome`, `PL.nome`, `from Modulo import`, nel dispatcher (`?Modulo.nome` in `Addons.xcu` e `Accelerators.xcu`) e nelle macro Basic. Le voci dentro commenti XML non contano.
5. `python3 scripts/check_pythonpath_safety.py` e `pytest --ignore=tests/test_brogliaccio_export.py tests/`.
6. Docstring: leggere a campione le prime 10 e tutte le funzioni lunghe o che sono wrapper.
7. Test nuovi: ogni test asserisce un valore restituito, almeno uno fa un roundtrip con dati sintetici, i `BUG NOTO` sono difetti veri da valutare.
8. Prova manuale in Calc per ciò che passa da UNO (contabilità, import, esportazione, barre degli strumenti).

## Consegna di patch

Se uno zip contiene un file che un agente sta ancora creando su `origin/dev`, fare `git pull` prima di estrarlo, altrimenti `both added`: procedura in `documentazione/LESSONS_GIT.md`.
