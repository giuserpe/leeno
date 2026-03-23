---
name: leeno-sviluppo
description: >
  Procedura per preparare l'ambiente di sviluppo di LeenO.
  Usare quando si inizia una sessione di sviluppo, si configura
  l'ambiente o si prepara il branch dev per lavorare.
---

# LeenO – Ambiente di Sviluppo

## Prerequisiti (una tantum)

1. Configurare Git per rispettare i fine riga esistenti:
   ```
   git config --global core.autocrlf false
   ```

2. Clonare il repository:
   ```
   git clone https://gitlab.com/giuserpe/leeno
   cd leeno
   git checkout dev
   ```

3. Creare il symlink all'estensione installata in LibreOffice:

   **Windows** (cmd come amministratore):
   ```
   cd %appdata%\LibreOffice\4\user\uno_packages\cache\uno_packages\
   cd [cartella_tmp]
   mv LeenO.oxt LeenO.oxt.old
   mklink /D LeenO.oxt W:\_dwg\ULTIMUSFREE\_SRC\leeno\src\Ultimus.oxt
   ```
   Oppure avviare `W:\_dwg\ULTIMUSFREE\LeenO.bat` come amministratore.

   **Linux**:
   ```
   cd $HOME/.config/libreoffice/4/user/uno_packages/cache/uno_packages/[cartella_tmp]
   mv LeenO.oxt LeenO.oxt.old
   ln -s /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt LeenO.oxt
   ```

## Procedura per ogni sessione di sviluppo

1. Passare al branch `dev`:
   ```
   git checkout dev
   ```

2. Cancellare la cache Python (se presente):
   ```
   rm -rf src/Ultimus.oxt/python/pythonpath/__pycache__
   ```

3. Aggiornare la versione in `src/Ultimus.oxt/python/pythonpath/LeenoGlobals.py`:
   - **`Lmajor`** → incrementare per INCOMPATIBILITÀ
   - **`Lminor`** → incrementare per NUOVE FUNZIONALITÀ
   - **`Lsubv`** → incrementare per CORREZIONE BUG

   > Questi valori influiscono su `su_apertura_doc` e `aggiorniamoli`.

4. Nella funzione `make_pack()` in `pyleeno.py`:
   - **Attivare** la chiamata a `description_upd()`

5. Pacchettizzare con **CTRL-SHIFT-G** in LibreOffice

6. Testare il pacchetto generato

## File di riferimento

| File | Ruolo |
|------|-------|
| `src/Ultimus.oxt/python/pythonpath/LeenoGlobals.py` | Variabili di versione |
| `src/Ultimus.oxt/python/pythonpath/pyleeno.py` | Funzione `make_pack()` |
| `src/Ultimus.oxt/description.xml` | Metadata estensione |
| `src/Ultimus.oxt/leeno_version_code` | Codice versione |
