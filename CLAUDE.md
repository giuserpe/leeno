# CLAUDE.md per LeenO

@AGENTS.md

## Specifico per Claude Code

- Lavora solo sul branch `dev`; `master` non si tocca senza istruzione esplicita.
- Le skill del progetto stanno in `.agent/skills/<nome>/SKILL.md` (cartella singolare, non `.claude/skills/`): quando il task le richiama (`git-commit-ita`, `leeno-aggiorna-manuale`, `leeno-icone-toolbar`, `leeno-sviluppo`, ...) leggi il relativo `SKILL.md` da lì prima di procedere.
- Nessun commit o push automatico se non richiesto: consegna lo zip con i comandi PowerShell come descritto in AGENTS.md e proponi il messaggio di commit senza eseguirlo.
- Prima di consegnare esegui la sezione "Verifica prima della consegna" di AGENTS.md.
- Niente emoji e niente trattini lunghi in risposte, commit e documenti.
