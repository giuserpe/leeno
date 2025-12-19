'''
Modulo di debug per LeenO
'''
import os
import sys
import pyleeno as PL
import Dialogs



def aggiorna_configurazione_leeno():
    '''Rigenera Addons.xcu e registrymodifications.xcu senza reinstallare (Win/Linux/Mac)'''

    # 1. Ottieni il percorso di LeenO installato
    leeno_path = PL.LeenO_path()

    # 2. Determina i percorsi in base al sistema operativo
    if sys.platform == "win32":
        user_config = os.path.expandvars(r"%APPDATA%\LibreOffice\4\user")
        source_addons = r"W:\_dwg\ULTIMUSFREE\_SRC\leeno\src\Ultimus.oxt\Addons.xcu"
        sistema = "Windows"
    elif sys.platform == "darwin":  # macOS
        user_config = os.path.expanduser("~/Library/Application Support/LibreOffice/4/user")
        source_addons = os.path.expanduser("~/path/to/dev/Ultimus.oxt/Addons.xcu")  # Modifica questo path
        sistema = "macOS"
    else:  # Linux (linux, linux2)
        user_config = os.path.expanduser("~/.config/libreoffice/4/user")
        source_addons = os.path.expanduser("~/path/to/dev/Ultimus.oxt/Addons.xcu")  # Modifica questo path
        sistema = "Linux"

    # 3. Trova e aggiorna Addons.xcu nella cache di configurazione
    config_base = os.path.join(user_config,
        "uno_packages/cache/registry/com.sun.star.comp.deployment.configuration.PackageRegistryBackend")

    if not os.path.exists(config_base):
        Dialogs.Info(Title="Errore", Text=f"Directory di configurazione non trovata:\n{config_base}")
        return False

    target_file = None
    for root, dirs, files in os.walk(config_base):
        if "Addons.xcu" in files:
            target_file = os.path.join(root, "Addons.xcu")
            break

    if not target_file:
        Dialogs.Info(Title="Errore", Text="Addons.xcu non trovato nella cache!")
        return False

    # 4. Verifica che il file sorgente esista
    if not os.path.exists(source_addons):
        Dialogs.Info(Title="Errore",
                     Text=f"File sorgente non trovato:\n{source_addons}\n\nModifica il path nel codice per {sistema}")
        return False

    # 5. Leggi e processa Addons.xcu
    try:
        with open(source_addons, 'r', encoding='utf-8') as f:
            content = f.read()

        # Sostituisci %origin% con il path reale
        content = content.replace('%origin%', leeno_path)

        # 6. Scrivi il file Addons.xcu processato
        with open(target_file, 'w', encoding='utf-8') as f:
            f.write(content)

        risultato_addons = "✓ Addons.xcu aggiornato"
    except Exception as e:
        Dialogs.Info(Title="Errore", Text=f"Errore nell'aggiornamento di Addons.xcu:\n{str(e)}")
        return False

    # 7. Cancella registrymodifications.xcu per aggiornare il dispatcher
    registry_file = os.path.join(user_config, "registrymodifications.xcu")

    try:
        if os.path.exists(registry_file):
            os.remove(registry_file)
            risultato_registry = "✓ registrymodifications.xcu eliminato"
        else:
            risultato_registry = "⚠ registrymodifications.xcu non trovato"
    except Exception as e:
        risultato_registry = f"✗ Errore cancellazione registrymodifications.xcu:\n{str(e)}"

    # 8. Messaggio finale
    msg = f"{risultato_addons}\n{risultato_registry}\n\nSistema: {sistema}"
    Dialogs.Info(Title="Configurazione di LeeenO Aggiornata",
                 Text=msg + "\n\nRiavvia LibreOffice per applicare le modifiche")
    return True
