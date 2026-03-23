"""
Utility per gestire undo/redo nelle operazioni LeenO
"""
import functools
import traceback


def with_undo(description=None, auto_description=True):
    """
    Decorator che aggiunge supporto undo a una funzione.

    Args:
        description (str, optional): Descrizione personalizzata dell'azione.
                                     Se None e auto_description=True, usa il nome della funzione.
        auto_description (bool): Se True, genera automaticamente la descrizione dal nome funzione.

    Usage:
        @with_undo
        def mia_funzione():
            # codice...

        @with_undo()
        def altra_funzione():
            # codice...

        @with_undo("Operazione personalizzata")
        def terza_funzione():
            # codice...
    """
    # Caso 1: @with_undo (senza parentesi) - description è la funzione
    if callable(description):
        func = description
        actual_description = func.__name__.replace('_', ' ').title()
        return _create_wrapper(func, actual_description)

    # Caso 2: @with_undo() o @with_undo("desc") - description è una stringa o None
    def decorator(func):
        # Determina la descrizione
        if description:
            actual_description = description
        elif auto_description:
            actual_description = func.__name__.replace('_', ' ').title()
        else:
            actual_description = "Operazione LeenO"

        return _create_wrapper(func, actual_description)

    return decorator


def _create_wrapper(func, undo_description):
    """
    Crea il wrapper effettivo per la funzione con supporto undo.
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # Import locale per evitare dipendenze circolari
        try:
            import LeenoUtils
            oDoc = LeenoUtils.getDocument()
            undoManager = oDoc.UndoManager
        except Exception:
            # Se non riesci ad ottenere l'undoManager, esegui comunque la funzione
            return func(*args, **kwargs)

        # Inizia contesto undo
        undoManager.enterUndoContext(undo_description)

        try:
            # Esegui la funzione
            result = func(*args, **kwargs)
            # Chiudi contesto undo in caso di successo
            undoManager.leaveUndoContext()
            return result

        except Exception as e:
            # In caso di errore, chiudi comunque il contesto
            try:
                undoManager.leaveUndoContext()
            except Exception:
                pass
            # Rilancia l'eccezione originale
            raise e

    return wrapper


def with_undo_batch(description="Operazioni multiple"):
    """
    Decorator per raggruppare multiple operazioni in un singolo undo.
    Utile quando una funzione chiama altre funzioni che hanno già @with_undo.

    Usage:
        @with_undo_batch("Importa dati complessi")
        def importa_tutto():
            importa_parte1()  # ha già @with_undo
            importa_parte2()  # ha già @with_undo
            # Risultato: un solo undo invece di due
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                import LeenoUtils
                oDoc = LeenoUtils.getDocument()
                undoManager = oDoc.UndoManager
            except Exception:
                return func(*args, **kwargs)

            # Blocca gli undo interni
            undoManager.lock()

            try:
                # Esegui le operazioni
                result = func(*args, **kwargs)

                # Sblocca e crea un singolo undo per tutto
                undoManager.unlock()
                undoManager.enterUndoContext(description)
                undoManager.leaveUndoContext()

                return result

            except Exception as e:
                # Sblocca in caso di errore
                if undoManager.isLocked():
                    undoManager.unlock()
                raise e

        return wrapper
    return decorator


def no_undo(func):
    """
    Decorator per disabilitare temporaneamente l'undo durante una funzione.
    Utile per operazioni di sola lettura o query.

    Usage:
        @no_undo
        def leggi_dati():
            # nessun undo creato
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            import LeenoUtils
            oDoc = LeenoUtils.getDocument()
            undoManager = oDoc.UndoManager
            was_locked = undoManager.isLocked()

            if not was_locked:
                undoManager.lock()

            try:
                result = func(*args, **kwargs)
                return result
            finally:
                if not was_locked:
                    undoManager.unlock()

        except Exception:
            # Se non riesci ad ottenere l'undoManager, esegui comunque
            return func(*args, **kwargs)

    return wrapper


class UndoContext:
    """
    Context manager per gestire undo in blocchi with.

    Usage:
        with UndoContext("Operazione complessa"):
            # codice che modifica il documento
            # tutto sarà in un singolo undo
    """
    def __init__(self, description="Operazione"):
        self.description = description
        self.undoManager = None

    def __enter__(self):
        try:
            import LeenoUtils
            oDoc = LeenoUtils.getDocument()
            self.undoManager = oDoc.UndoManager
            self.undoManager.enterUndoContext(self.description)
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.undoManager:
            try:
                self.undoManager.leaveUndoContext()
            except Exception:
                pass
        return False  # Non sopprimere eccezioni


def clear_undo_history():
    """
    Svuota completamente la storia undo/redo.
    Usa con cautela!
    """
    try:
        import LeenoUtils
        oDoc = LeenoUtils.getDocument()
        oDoc.UndoManager.clear()
    except Exception as e:
        import LeenoDialogs as DLG
        DLG.chi(f"Impossibile svuotare undo history: {e}")


########################################################################


def analizza_performance_log():
    """
    Analizza il file di log delle performance e mostra statistiche.
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
            print("Nessun log di performance trovato")
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
                            if "ms" in time_str:
                                ms = float(time_str.replace("ms", ""))
                            elif "s" in time_str:
                                ms = float(time_str.replace("s", "")) * 1000
                            else:
                                continue

                            stats[func_name].append(ms)

        # Mostra statistiche
        print("\n=== STATISTICHE PERFORMANCE ===\n")
        for func_name, times in sorted(stats.items()):
            avg_time = sum(times) / len(times)
            min_time = min(times)
            max_time = max(times)
            count = len(times)

            print(f"{func_name}:")
            print(f"  Chiamate: {count}")
            print(f"  Media: {avg_time:.2f} ms")
            print(f"  Min: {min_time:.2f} ms")
            print(f"  Max: {max_time:.2f} ms")
            print()

    except Exception as e:
        print(f"Errore nell'analisi del log: {e}")
