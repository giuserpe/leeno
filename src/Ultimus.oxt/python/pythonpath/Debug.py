#!/usr/bin/env python3
# -*- Mode: Python; coding: utf-8; indent-tabs-mode: nil; tab-width: 4 -*-
########################################################################
'''
Modulo di debug per LeenO
'''
import os
import sys
import pyleeno as PL
import Dialogs
import LeenoUtils
import LeenoDialogs as DLG


# def aggiorna_configurazione_leeno():
#     '''Rigenera Addons.xcu e registrymodifications.xcu senza reinstallare (Win/Linux/Mac)'''

#     # 1. Ottieni il percorso di LeenO installato
#     leeno_path = PL.LeenO_path()

#     # 2. Determina i percorsi in base al sistema operativo
#     if sys.platform == "win32":
#         user_config = os.path.expandvars(r"%APPDATA%\LibreOffice\4\user")
#         source_addons = r"W:\_dwg\ULTIMUSFREE\_SRC\leeno\src\Ultimus.oxt\Addons.xcu"
#         sistema = "Windows"
#     elif sys.platform == "darwin":  # macOS
#         user_config = os.path.expanduser("~/Library/Application Support/LibreOffice/4/user")
#         source_addons = os.path.expanduser("~/path/to/dev/Ultimus.oxt/Addons.xcu")  # Modifica questo path
#         sistema = "macOS"
#     else:  # Linux (linux, linux2)
#         user_config = os.path.expanduser("~/.config/libreoffice/4/user")
#         source_addons = os.path.expanduser("~/path/to/dev/Ultimus.oxt/Addons.xcu")  # Modifica questo path
#         sistema = "Linux"

#     # 3. Trova e aggiorna Addons.xcu nella cache di configurazione
#     config_base = os.path.join(user_config,
#         "uno_packages/cache/registry/com.sun.star.comp.deployment.configuration.PackageRegistryBackend")

#     if not os.path.exists(config_base):
#         Dialogs.Info(Title="Errore", Text=f"Directory di configurazione non trovata:\n{config_base}")
#         return False

#     target_file = None
#     for root, dirs, files in os.walk(config_base):
#         if "Addons.xcu" in files:
#             target_file = os.path.join(root, "Addons.xcu")
#             break

#     if not target_file:
#         Dialogs.Info(Title="Errore", Text="Addons.xcu non trovato nella cache!")
#         return False

#     # 4. Verifica che il file sorgente esista
#     if not os.path.exists(source_addons):
#         Dialogs.Info(Title="Errore",
#                      Text=f"File sorgente non trovato:\n{source_addons}\n\nModifica il path nel codice per {sistema}")
#         return False

#     # 5. Leggi e processa Addons.xcu
#     try:
#         with open(source_addons, 'r', encoding='utf-8') as f:
#             content = f.read()

#         # Sostituisci %origin% con il path reale
#         content = content.replace('%origin%', leeno_path)

#         # 6. Scrivi il file Addons.xcu processato
#         with open(target_file, 'w', encoding='utf-8') as f:
#             f.write(content)

#         risultato_addons = "✓ Addons.xcu aggiornato"
#     except Exception as e:
#         Dialogs.Info(Title="Errore", Text=f"Errore nell'aggiornamento di Addons.xcu:\n{str(e)}")
#         return False

#     # 7. Cancella registrymodifications.xcu per aggiornare il dispatcher
#     registry_file = os.path.join(user_config, "registrymodifications.xcu")

#     try:
#         if os.path.exists(registry_file):
#             os.remove(registry_file)
#             risultato_registry = "✓ registrymodifications.xcu eliminato"
#         else:
#             risultato_registry = "⚠ registrymodifications.xcu non trovato"
#     except Exception as e:
#         risultato_registry = f"✗ Errore cancellazione registrymodifications.xcu:\n{str(e)}"

#     # 8. Messaggio finale
#     msg = f"{risultato_addons}\n{risultato_registry}\n\nSistema: {sistema}"
#     Dialogs.Info(Title="Configurazione di LeeenO Aggiornata",
#                  Text=msg + "\n\nRiavvia LibreOffice per applicare le modifiche")
    # return True


import uno

import uno
import os
import sys

import uno
import os
import sys

def force_restart():
    """Tenta una chiusura pulita di LibreOffice tramite API UNO"""
    try:
        # Ottieni il contesto (funziona sia in macro che in script Python-UNO)
        local_context = uno.getComponentContext()
        smgr = local_context.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", local_context)
        
        # terminate() restituisce True se la chiusura è stata accettata
        chiuso = desktop.terminate()
        return chiuso
    except Exception as e:
        # In caso di errore (es. servizio non disponibile)
        Dialogs.Info(Title="Errore chiusura", Text=f"Impossibile usare terminate():\n{str(e)}")
        return False


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
        source_addons = os.path.expanduser("~/path/to/dev/Ultimus.oxt/Addons.xcu")  # ← MODIFICA
        sistema = "macOS"
    else:  # Linux
        user_config = os.path.expanduser("~/.config/libreoffice/4/user")
        source_addons = os.path.expanduser("~/path/to/dev/Ultimus.oxt/Addons.xcu")  # ← MODIFICA
        sistema = "Linux"

    # 3. Trova Addons.xcu nella cache
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

    # 4. Verifica sorgente
    if not os.path.exists(source_addons):
        Dialogs.Info(Title="Errore",
                     Text=f"File sorgente non trovato:\n{source_addons}\n\nModifica il path nel codice per {sistema}")
        return False

    # 5. Aggiorna Addons.xcu
    try:
        with open(source_addons, 'r', encoding='utf-8') as f:
            content = f.read()
        content = content.replace('%origin%', leeno_path)
        
        with open(target_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        risultato_addons = "✓ Addons.xcu aggiornato"
    except Exception as e:
        Dialogs.Info(Title="Errore", Text=f"Errore nell'aggiornamento di Addons.xcu:\n{str(e)}")
        return False

    # 6. Elimina registrymodifications.xcu
    registry_file = os.path.join(user_config, "registrymodifications.xcu")
    try:
        if os.path.exists(registry_file):
            os.remove(registry_file)
            risultato_registry = "✓ registrymodifications.xcu eliminato"
        else:
            risultato_registry = "⚠ registrymodifications.xcu non trovato (non grave)"
    except Exception as e:
        risultato_registry = f"✗ Errore cancellazione registrymodifications.xcu:\n{str(e)}"

    # 7. Messaggio di riepilogo + richiesta riavvio
    msg = (
        f"{risultato_addons}\n"
        f"{risultato_registry}\n\n"
        f"Sistema: {sistema}\n\n"
        "Per applicare le modifiche è necessario riavviare LibreOffice.\n"
        "Vuoi chiuderlo adesso?"
    )

    risposta = Dialogs.YesNoDialog(
        Title="Chiusura LibreOffice?",
        Text=msg
    )

    if risposta:  # Sì → tenta chiusura pulita
        Dialogs.Info(
            Title="Chiusura in corso",
            Text="Sto chiudendo LibreOffice...\n\n"
                 "Riaprilo manualmente per applicare le modifiche a LeenO.\n"
                 "(doppio clic su un file .ods o su soffice.exe / LibreOffice.app / libreoffice)"
        )

        chiuso = force_restart()

        if not chiuso:
            Dialogs.Info(
                Title="Chiusura non completata",
                Text="La chiusura automatica è stata bloccata.\n"
                     "Probabilmente hai documenti aperti con modifiche non salvate.\n\n"
                     "Salva tutto, chiudi LibreOffice manualmente e riaprilo."
            )

    else:
        Dialogs.Info(
            Title="Modifiche completate",
            Text="Ok, hai scelto di non chiudere ora.\n"
                 "Ricorda di chiudere e riaprire LibreOffice quando vuoi applicare le modifiche."
        )

    return True



########################################################################
# funzioni per misurare performance delle funzioni LeenO
# elenco 3 versioni:
# 1) measure_time: decorator completo con opzioni per popup, log su file,
#    soglia di tempo per popup, log in console (solo per sviluppatore)
# 2) measure_time_simple: versione semplificata che logga solo su file
# 3) PerformanceMonitor: context manager per misurare blocchi di codice
########################################################################

"""
Utility per misurare performance delle funzioni LeenO
"""
import functools
import time
from datetime import datetime


def _format_time(seconds):
    """Formatta il tempo in modo leggibile."""
    if seconds < 0.001:
        return f"{seconds * 1000000:.0f} μs"
    elif seconds < 1:
        return f"{seconds * 1000:.2f} ms"
    elif seconds < 60:
        return f"{seconds:.2f} s"
    else:
        minutes = int(seconds // 60)
        secs = seconds % 60
        return f"{minutes}m {secs:.2f}s"


def _log_to_file(func_name, module_name, start_datetime, elapsed_time, success, error_msg):
    """Scrive le informazioni nel file di log."""
    try:
        import os
        import uno

        # Ottieni il path dell'estensione
        pir = uno.getComponentContext().getValueByName(
            '/singletons/com.sun.star.deployment.PackageInformationProvider'
        )
        expath_url = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
        expath = uno.fileUrlToSystemPath(expath_url)

        log_dir = os.path.join(expath, "pythonpath")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "leeno_performance.log")

        with open(log_path, "a", encoding="utf-8") as logfile:
            timestamp = start_datetime.strftime('%Y-%m-%d %H:%M:%S')
            status = "SUCCESS" if success else "ERROR"
            time_str = _format_time(elapsed_time)

            logfile.write(f"[{timestamp}] [{status}] {module_name}.{func_name}: {time_str}")
            if error_msg:
                logfile.write(f" - Error: {error_msg}")
            logfile.write("\n")

    except Exception:
        pass  # Silenzioso se non riesce a scrivere il log


def _log_to_console(func_name, time_str, success):
    """Scrive in console (solo per debug in sviluppo)."""
    try:
        import os
        import uno

        # Solo per l'utente giuserpe (sviluppatore)
        user = os.environ.get("USERNAME", "").lower()
        if "giuserpe" in user:
            pir = uno.getComponentContext().getValueByName(
                '/singletons/com.sun.star.deployment.PackageInformationProvider'
            )
            expath_url = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
            expath = uno.fileUrlToSystemPath(expath_url)

            console_log = os.path.join(expath, "pythonpath", "leeno_console.log")

            status = "✓" if success else "✗"
            msg = f"[{status}] {func_name}: {time_str}\n"

            with open(console_log, "a", encoding="utf-8") as f:
                f.write(msg)
    except Exception:
        pass


def _show_popup(func_name, time_str, success, error_msg):
    """Mostra un popup con le informazioni sul tempo."""
    try:
        import LeenoDialogs as DLG
        status = "✓ Completata" if success else "✗ Errore"
        msg = f"Funzione: {func_name}\nTempo: {time_str}\nStato: {status}"
        if error_msg:
            msg += f"\nErrore: {error_msg}"
        DLG.chi(msg, OFF=True)
    except Exception:
        pass


def measure_time(show_popup=True, log_to_file=True, threshold_seconds=None, console_log=False):
    """
    Decorator per misurare il tempo di esecuzione di una funzione.

    Args:
        show_popup (bool): Se True, mostra un popup con il tempo di esecuzione
        log_to_file (bool): Se True, scrive i tempi in un file di log
        threshold_seconds (float): Se specificato, mostra popup solo se il tempo supera questa soglia
        console_log (bool): Se True, scrive anche in console (solo per sviluppatore)

    Usage:
        @measure_time()
        def funzione_lenta():
            # codice...

        @measure_time(show_popup=True)
        def funzione_importante():
            # mostra sempre il tempo

        @measure_time(threshold_seconds=1.0)
        def funzione_critica():
            # mostra popup solo se impiega più di 1 secondo
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Timestamp inizio
            start_time = time.time()
            start_datetime = datetime.now()

            # Informazioni sulla funzione
            func_name = func.__name__
            module_name = func.__module__

            success = True
            error_msg = None
            result = None

            try:
                # Esegui la funzione
                result = func(*args, **kwargs)

            except Exception as e:
                success = False
                error_msg = str(e)
                raise  # Rilancia l'eccezione

            finally:
                # Timestamp fine
                end_time = time.time()
                elapsed_time = end_time - start_time

                # Formatta il tempo in modo leggibile
                time_str = _format_time(elapsed_time)

                # Log su file
                if log_to_file:
                    _log_to_file(
                        func_name, module_name, start_datetime,
                        elapsed_time, success, error_msg
                    )

                # Log in console (solo per sviluppatore)
                if console_log:
                    _log_to_console(func_name, time_str, success)

                # Mostra popup se richiesto o se supera la soglia
                should_show = show_popup or (
                    threshold_seconds is not None and elapsed_time > threshold_seconds
                )

                if should_show:
                    _show_popup(func_name, time_str, success, error_msg)

            return result

        return wrapper

    return decorator


def measure_time_simple(func):
    """
    Versione semplificata del decorator per misurare i tempi.
    Logga solo su file, nessun popup.

    Usage:
        @measure_time_simple
        def mia_funzione():
            # codice...
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start = time.time()
        start_datetime = datetime.now()

        try:
            result = func(*args, **kwargs)
            elapsed = time.time() - start

            # Log su file
            _log_to_file(
                func.__name__,
                func.__module__,
                start_datetime,
                elapsed,
                True,
                None
            )

            return result

        except Exception as e:
            elapsed = time.time() - start

            # Log su file con errore
            _log_to_file(
                func.__name__,
                func.__module__,
                start_datetime,
                elapsed,
                False,
                str(e)
            )

            raise

    return wrapper

class PerformanceMonitor:
    """
    Context manager per misurare blocchi di codice.

    Usage:
        with PerformanceMonitor("Caricamento dati"):
            carica_dati()
            elabora_dati()
            # Al termine logga: "Caricamento dati: 2.34 s"

        with PerformanceMonitor("Operazione critica", show_popup=True):
            operazione()
            # Mostra anche un popup
    """
    def __init__(self, name="Operazione", show_popup=False, log_to_file=True):
        self.name = name
        self.show_popup = show_popup
        self.log_to_file = log_to_file
        self.start_time = None
        self.start_datetime = None

    def __enter__(self):
        self.start_time = time.time()
        self.start_datetime = datetime.now()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        elapsed = time.time() - self.start_time
        time_str = _format_time(elapsed)

        success = exc_type is None
        error_msg = str(exc_val) if exc_val else None

        # Log su file
        if self.log_to_file:
            _log_to_file(
                self.name,
                "__context__",
                self.start_datetime,
                elapsed,
                success,
                error_msg
            )

        # Mostra popup se richiesto
        if self.show_popup:
            _show_popup(self.name, time_str, success, error_msg)

        return False  # Non sopprimere eccezioni


def mostra_statistiche_performance():
    """
    Analizza il file di log delle performance e mostra statistiche in un popup.
    """
    import os
    import uno
    from collections import defaultdict

    try:
        pir = uno.getComponentContext().getValueByName(
            '/singletons/com.sun.star.deployment.PackageInformationProvider'
        )
        expath_url = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
        expath = uno.fileUrlToSystemPath(expath_url)
        log_path = os.path.join(expath, "pythonpath", "leeno_performance.log")

        if not os.path.exists(log_path):
            import LeenoDialogs as DLG
            DLG.chi("Nessun log di performance trovato")
            return

        stats = defaultdict(list)

        with open(log_path, "r", encoding="utf-8") as f:
            for line in f:
                if "SUCCESS" in line or "ERROR" in line:
                    parts = line.split("] ")
                    if len(parts) >= 3:
                        func_info = parts[2].split(": ")
                        if len(func_info) >= 2:
                            func_name = func_info[0].strip()
                            time_str = func_info[1].split()[0]

                            # Converti in millisecondi
                            try:
                                if "ms" in time_str:
                                    ms = float(time_str.replace("ms", ""))
                                elif "s" in time_str and "m" not in time_str:
                                    ms = float(time_str.replace("s", "")) * 1000
                                else:
                                    continue

                                stats[func_name].append(ms)
                            except ValueError:
                                continue

        # Genera messaggio con statistiche
        msg = "=== STATISTICHE PERFORMANCE ===\n\n"

        # Ordina per tempo medio decrescente
        sorted_stats = sorted(
            stats.items(),
            key=lambda x: sum(x[1]) / len(x[1]),
            reverse=True
        )

        for func_name, times in sorted_stats[:20]:  # Top 20
            avg_time = sum(times) / len(times)
            min_time = min(times)
            max_time = max(times)
            count = len(times)

            msg += f"{func_name}:\n"
            msg += f"  Chiamate: {count}\n"
            msg += f"  Media: {avg_time:.2f} ms\n"
            msg += f"  Min: {min_time:.2f} ms\n"
            msg += f"  Max: {max_time:.2f} ms\n\n"

        import LeenoDialogs as DLG
        DLG.chi(msg)

    except Exception as e:
        import LeenoDialogs as DLG
        DLG.chi(f"Errore nell'analisi del log: {e}")


def pulisci_log_performance():
    """
    Svuota il file di log delle performance.
    """
    import os
    import uno

    try:
        pir = uno.getComponentContext().getValueByName(
            '/singletons/com.sun.star.deployment.PackageInformationProvider'
        )
        expath_url = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
        expath = uno.fileUrlToSystemPath(expath_url)
        log_path = os.path.join(expath, "pythonpath", "leeno_performance.log")

        if os.path.exists(log_path):
            os.remove(log_path)
            import LeenoDialogs as DLG
            DLG.chi("Log di performance eliminato con successo")
        else:
            import LeenoDialogs as DLG
            DLG.chi("Nessun log di performance da eliminare")

    except Exception as e:
        import LeenoDialogs as DLG
        DLG.chi(f"Errore nell'eliminazione del log: {e}")