---
name: leeno-rilascio
description: >
  Procedura completa per il rilascio di una nuova versione di LeenO.
  Usare quando si vuole pubblicare una release: preparazione contenuti,
  aggiornamento template, pacchetto OXT, Git tag, upload sito e comunicazione.
---

# LeenO – Rilascio Nuova Versione

## Fase A — Preparazione contenuti

1. **Generazione Articolo Novità**:
   Invocare la skill `leeno-articolo-novita` per analizzare i cambiamenti dall'ultimo tag stabile e produrre una bozza dell'articolo di presentazione orientata all'utente.
   - Alternativamente, comando rapido per il log: `git log vX.X.X..HEAD --pretty=format:"%s" --reverse`
2. **Note di Versione**: Riportare le novità salienti nel manuale (`documentazione/MANUALE_LeenO.fodt`).
3. **Esportazione Manuale**: Generare il PDF aggiornato (`MANUALE_LeenO.pdf`).
4. **Finalizzazione Articolo**: Rifinire la bozza generata al punto 1 per l'uso su blog e social.

## Fase B — Aggiornamento template .ODS

1. Cancellare eventuali fogli extra usati per prove
2. Controllare il contenuto del foglio **S1**
3. Cambiare le **proprietà** del `.ODS` (es. numero di versione)
4. Impostare **Visualizza > Normale**, **Evidenzia valori > OFF**
5. Controllare esistenza e correttezza macro su **Personalizza > Eventi**
6. Cancellare anagrafica generale e situazione contabile
7. Aggiornare `def adegua_tmpl` al numero del template
8. Sostituire il manuale in PDF nel template

## Fase C — Preparazione pacchetto OXT

1. Verificare la versione in `LeenoGlobals.py` (`Lmajor`, `Lminor`, `Lsubv`)
2. In `make_pack()`: **disattivare** `description_upd()`
3. Controllare/aggiornare versione in `description.xml`
4. Controllare/aggiornare versione in `leeno_version_code`
5. Aggiornare info in `pkg-desc/description.txt`
6. Rimuovere `def MENU_debug` (se presente)
7. Cancellare `MANUALE_LeenO.pdf` dalla dir sorgente (se presente)
8. Aggiungere `MANUALE_LeenO.pdf` aggiornato al pacchetto
9. Pacchettizzare con **CTRL-SHIFT-G**
10. **✅ Verificare**: assenza di `__pycache__` nel pacchetto OXT
11. **✅ Verificare**: versione corretta in `leeno_version_code` nel pacchetto
12. Rinominare il pacchetto come `LeenO-X.XX.X.oxt`

## Fase D — Verifica finale

1. Cancellare `leeno.conf`
2. Reinstallare l'OXT da zero
3. Verificare che le informazioni visulizzate siano corrette

## Fase E — Pubblicazione Git

1. Commit di tutti gli aggiornamenti su `dev`:
   ```
   git add .
   git commit -m "feat: rilascio versione X.XX.X"
   git push origin dev
   ```

2. Merge in master:
   ```
   git checkout master
   git merge dev
   ```

3. Creare il tag annotato:
   ```
   git tag -a vX.XX.X -m 'Release LeenO versione X.XX.X'
   ```

4. Push verso tutti i remoti:
   ```
   git push origin master
   git push origin --tags
   git push GH master
   git push GH --tags
   ```

> **Nota**: usare `git tag -a` (annotato) anziché `git tag -s` (firmato)
> se non è configurata una chiave GPG.

## Fase F — Pubblicazione sito e distribuzione

1. Copiare OXT e file di esempio in **INCUBATRICE**
2. Upload su [gestione file leeno.org](https://leeno.org/wp-admin/admin.php?page=wpfilebase_filebrowser)
   - Spostare la vecchia versione nella categoria **Archivio**
3. Aggiornare la versione in [LeenO.update.xml](https://leeno.org/LeenO.update.xml)
4. Eseguire [Sincronizza Filebase](https://leeno.org/wp-admin/admin.php?page=wpfilebase_manage&action=sync&no-ob=1)
5. Correggere la versione nella [pagina download](https://leeno.org/about-leeno/leeno/download/)

## Fase G — Comunicazione

1. Aggiornare [DOXYGEN](https://leeno.org/doxyLeenO/html/namespacepyleeno.html)
2. Pubblicare l'articolo (generato nella Fase A) su **Linkedin**, **openikos**, **X** e **Telegram**.
3. Aggiornare [LibreOffice Extensions](https://extensions.libreoffice.org/extensions/leeno-2).
4. Inviare **Newsletter** basata sullo stesso articolo.

---
> [!TIP]
> Durante la Fase A, puoi chiedere all'assistente: *"Usa la skill leeno-articolo-novita per preparare il post di rilascio"* per ottenere una bozza già pronta e orientata all'utente finale.
