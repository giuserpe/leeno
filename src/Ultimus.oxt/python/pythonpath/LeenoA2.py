# pyrefly: ignore [missing-import]
import uno
import os
from pathlib import Path

import LeenoUtils
import Dialogs

FOGLI_TARGET = ["COMPUTO", "VARIANTE", "CONTABILITA"]

def fmt_euro(val):
    try:
        n = float(val)
        s = f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"€ {s}"
    except (ValueError, TypeError):
        return str(val) if val else "—"

def MENU_importi():
    ctx = LeenoUtils.getComponentContext()
    smgr = ctx.ServiceManager
    
    # 1. Seleziona cartella
    picker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", ctx)
    picker.setTitle("Seleziona la cartella con i file ODS")
    
    if picker.execute() != 1:
        return
        
    folder_url = picker.getDirectory()
    folder_path = uno.fileUrlToSystemPath(folder_url)
    cartella = Path(folder_path)
    
    ods_files = sorted(cartella.glob("*.ods"))
    
    if not ods_files:
        Dialogs.NotifyDialog(Title="Nessun file", Text=f"Nessun file ODS trovato in:\n{cartella.resolve()}")
        return

    # Inizializza testo
    righe = [
        f"{cartella.resolve()}",
        f"File ODS: {len(ods_files)}",
        "=" * 40,
        ""
    ]
    
    desktop = LeenoUtils.getDesktop()
    # pyrefly: ignore [missing-import]
    from com.sun.star.beans import PropertyValue
    props = (
        PropertyValue("Hidden", 0, True, 0),
        PropertyValue("ReadOnly", 0, True, 0),
        PropertyValue("UpdateDocMode", 0, 0, 0) # Non aggiorna i link esterni per velocizzare
    )

    for path in ods_files:
        righe.append(f"📄 {path.name}")
        url = uno.systemPathToFileUrl(str(path))
        
        try:
            # Apri il documento in background
            doc = desktop.loadComponentFromURL(url, "_blank", 0, props)
            if not doc:
                righe.append("   ❌ Impossibile aprire il file")
                righe.append("")
                continue
                
            sheets = doc.getSheets()
            trovato = False
            
            for target in FOGLI_TARGET:
                if sheets.hasByName(target):
                    trovato = True
                    sheet = sheets.getByName(target)
                    try:
                        # Leggi A2 (colonna 0, riga 1)
                        cell = sheet.getCellByPosition(0, 1)
                        val = cell.Value
                        # Se la cella è vuota o testo puro, tentiamo la conversione oppure usiamo la stringa
                        if cell.Type.value == 'TEXT':
                            val_str = cell.String
                        else:
                            val_str = fmt_euro(val)
                            
                        # Allinea il testo
                        importo = val_str.rjust(15,".")
                        target = target.ljust(12,".")
                        righe.append(f"   {target}{importo}")
                    except Exception as e:
                        righe.append(f"   {target}  [errore lettura]")
            
            if not trovato:
                righe.append("     (nessuno dei fogli target trovato)")
                
            doc.close(True)
            
        except Exception as e:
            righe.append(f"   ❌ {e}")
            
        righe.append("")

    # Mostra i risultati in una dialog nativa
    testo_completo = "\n".join(righe)
    # Dialogs.NotifyDialog(Title="Importi COMPUTO / VARIANTE / CONTABILITA", Text=testo_completo, FontName="Liberation Mono", FontWeight=100)

    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance('com.sun.star.awt.DialogProvider')

    oDlg = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.Dlg_importi?language=Basic&location=application"
    )

    listbox = oDlg.getControl('ListBox1')
    oDlg.Title = righe[0]
    listbox.addItems(righe[4:], 0)

    listbox.Model.FontName = "Liberation Mono"
    # listbox.Model.FontWeight = 100

    # ods_url = uno.systemPathToFileUrl(str(cartella.resolve()))
    
    props = (
        PropertyValue("Hidden", 0, False, 0),
        PropertyValue("ReadOnly", 0, False, 0),
    )


    oDlg.endExecute()
    oDlg.execute()

    nome = listbox.SelectedItem
    folder_path = Path(righe[0])
    ods_path = folder_path / nome[2:]
    ods_url = uno.systemPathToFileUrl(str(ods_path))
    try:
        doc2 = desktop.loadComponentFromURL(ods_url, "_blank", 0, props)
    except Exception as e:
        # Dialogs.NotifyDialog(Title="Errore", Text=f"Errore apertura file:\n{e}")
        pass
