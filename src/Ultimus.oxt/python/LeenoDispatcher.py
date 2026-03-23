import sys
import os
import inspect
import importlib
import traceback
import uno
import unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from datetime import datetime

def msgbox(*, Title='Errore interno', Message=''):
    """ Create message box in LibreOffice """
    try:
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstance('com.sun.star.awt.Toolkit')
        parent = toolkit.getDesktopWindow()
        buttons = MSG_BUTTONS.BUTTONS_OK
        type_msg = 'errorbox'
        mb = toolkit.createMessageBox(parent, type_msg, buttons, Title, str(Message))
        return mb.execute()
    except:
        print("Errore nella creazione del msgbox:", Message)
        return -1

def loVersion():
    aConfigProvider = uno.getComponentContext().ServiceManager.createInstance(
        "com.sun.star.configuration.ConfigurationProvider")
    arg = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    arg.Name = "nodepath"
    arg.Value = '/org.openoffice.Setup/Product'
    return aConfigProvider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess", (arg, )
    ).ooSetupVersionAboutBox

def fixPythonPath():
    myPath = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
    myPath = os.path.join(myPath, "pythonpath")
    if myPath not in sys.path:
        sys.path.append(myPath)

def reloadLeenoModules():
    myPath = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
    myPath = os.path.join(myPath, "pythonpath")
    if not os.path.exists(myPath):
        return
    pythonFiles = [f[:-3] for f in os.listdir(myPath) if os.path.isfile(os.path.join(myPath, f)) and f.endswith(".py")]
    for f in pythonFiles:
        try:
            module = importlib.import_module(f)
            importlib.reload(module)
        except Exception:
            print(f"Errore ricaricando il modulo: {f}")

def cerca_path_valido():
    """
    Cerca il percorso di un editor di codice installato sul sistema.
    Priorità: Antigravity > VS Code
    Supporta Windows, Linux e macOS.

    Returns:
        str: Percorso completo dell'editor trovato

    Raises:
        FileNotFoundError: Se nessun editor viene trovato
    """
    if 'giuserpe' in os.getlogin():
        import platform
        system = platform.system()

        possible_paths = []

        # === ANTIGRAVITY (Priorità 1) ===
        if system == "Windows":
            possible_paths.extend([
                os.path.expanduser("~\\AppData\\Local\\Programs\\Antigravity\\Antigravity.exe"),
                "C:\\Program Files\\Antigravity\\Antigravity.exe",
                "C:\\Program Files (x86)\\Antigravity\\Antigravity.exe",
                "C:\\Users\\TEST\\AppData\\Local\\Programs\\Antigravity\\Antigravity.exe",
            ])
        elif system == "Linux":
            possible_paths.extend([
                "/usr/bin/antigravity",
                "/usr/local/bin/antigravity",
                os.path.expanduser("~/.local/bin/antigravity"),
            ])
        elif system == "Darwin":  # macOS
            possible_paths.extend([
                "/Applications/Antigravity.app/Contents/MacOS/Antigravity",
                os.path.expanduser("~/Applications/Antigravity.app/Contents/MacOS/Antigravity"),
            ])

        # === VS CODE (Fallback) ===
        if system == "Windows":
            possible_paths.extend([
                os.path.expanduser("~\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"),
                "C:\\Program Files\\Microsoft VS Code\\Code.exe",
                "C:\\Program Files (x86)\\Microsoft VS Code\\Code.exe",
                "C:\\Users\\giuserpe\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
                "C:\\Users\\DELL\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            ])
        elif system == "Linux":
            possible_paths.extend([
                "/usr/bin/code",
                "/usr/local/bin/code",
                "/snap/bin/code",
                os.path.expanduser("~/.local/bin/code"),
            ])
        elif system == "Darwin":  # macOS
            possible_paths.extend([
                "/Applications/Visual Studio Code.app/Contents/Resources/app/bin/code",
                "/usr/local/bin/code",
            ])

        editor_path = None
        for path in possible_paths:
            if os.path.exists(path):
                editor_path = path
                break

        if editor_path is None:
            raise FileNotFoundError(
                f"Impossibile trovare un editor (Antigravity o VS Code) su {system}. "
                "Assicurati che almeno uno sia installato."
            )
        return editor_path

def dest():
    '''Definisce il percorso di destinazione basato sul sistema operativo'''

    # Per sistema Windows
    if os.name == 'nt':
        if not os.path.exists('w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/'):
            try:
                os.makedirs(
                    os.getenv("HOMEPATH") + '\\' + src_oxt +
                    '\\leeno\\src\\Ultimus.oxt\\')
            except FileExistsError:
                pass
            return os.getenv("HOMEDRIVE") + os.getenv(
                "HOMEPATH") + '\\' + src_oxt + '\\leeno\\src\\Ultimus.oxt\\'
        else:
            return 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'

    # Per sistemi Linux o macOS
    else:
        dest_path = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt/python/pythonpath'
        if not os.path.exists(dest_path):
            try:
                dest_path = os.getenv(
                    "HOME") + '/' + src_oxt + '/leeno/src/Ultimus.oxt/'
                os.makedirs(dest_path)
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/leeno/bin/')
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/_SRC/OXT')
            except FileExistsError:
                pass
        return dest_path


def handle_exception(e):
    ''' Gestisce gli errori e mostra un messaggio diagnostico '''
    try:
        # Recupero informazioni sul pacchetto installato (Provider)
        pir = uno.getComponentContext().getValueByName(
            '/singletons/com.sun.star.deployment.PackageInformationProvider')
        expath_url = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
        expath = uno.fileUrlToSystemPath(expath_url)

        # Lettura versione interna dal file code
        code_file = os.path.join(expath, 'leeno_version_code')
        version_line = ''
        if os.path.exists(code_file):
            with open(code_file, 'r', encoding='utf-8') as f:
                version_line = f.readline().strip()

        # Messaggio diagnostico base
        msg = (
            f"OS: {sys.platform} / LibreOffice-{loVersion()} / {version_line}\n\n"
            f"Errore: {str(e)}\n\n"
        )

        user = os.environ.get("USERNAME", "").lower()
        sysinfo = sys.exc_info()
        tb = sysinfo[2]

        if tb:
            tbInfo = traceback.extract_tb(tb)[-1]
            full_path_provider = tbInfo.filename # Percorso nella cache di LO
            line = tbInfo.lineno

            msg += (
                f"File: '{os.path.basename(full_path_provider)}'\n"
                f"Line: '{line}'\n"
                f"Function: '{tbInfo.name}'\n"
            )

            # ==========================================================
            # DEBUG AUTOMATICO PER GIUSEPPE (giuserpe)
            # ==========================================================
            if "giuserpe" in user:
                try:
                    import subprocess

                    # Ottengo la radice dei sorgenti (workspace)
                    src_root = dest()

                    # Calcolo il percorso relativo del file rispetto all'estensione installata
                    # e lo rimappo sul percorso dei sorgenti in W:
                    rel_path = os.path.relpath(full_path_provider, expath)
                    target_file = os.path.join(src_root, rel_path)

                    # Se il file nei sorgenti non esiste, fallback sul file del provider
                    if not os.path.exists(target_file):
                        target_file = full_path_provider

                    editor_path = cerca_path_valido() #
                    if os.path.exists(editor_path):
                        # COMANDO: [Editor] [Cartella Workspace] -g [File]:[Riga]
                        # Questo apre la cartella W:\... come workspace e il file sorgente
                        subprocess.Popen([editor_path, src_root, "-g", f"{target_file}:{line}"])
                    else:
                        msg += "\n(Editor non trovato per l'apertura automatica)\n"

                except Exception as dbg_err:
                    msg += f"\n(Errore automazione debug: {dbg_err})\n"
            # ==========================================================

        # Generazione Backtrace per il log
        msg += "+--" * 20 + "\nBACKTRACE:\n"
        for bk in traceback.extract_tb(tb):
            msg += f"File: {os.path.basename(bk.filename)}, Line: {bk.lineno}, Function: {bk.name}\n"

        # Scrittura del log nel profilo dell'estensione
        log_dir = os.path.join(expath, "pythonpath")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "leeno_error.log")
        with open(log_path, "a", encoding="utf-8") as logfile:
            logfile.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
            logfile.write(msg)
            logfile.write("-" * 60 + "\n\n")

        # Mostra dettagli estesi nel messaggio solo per giuserpe
        if "giuserpe" in user:
            msg += "+--" * 20 + "\n" + traceback.format_exc()

        msgbox(Title="Errore interno", Message=msg) #

    except Exception as internal:
        fallback = f"Errore nel gestore errori:\n{str(internal)}\n\n{traceback.format_exc()}"
        try:
            msgbox(Title="Errore gestore", Message=fallback)
        except:
            print(fallback)


# Dispatcher class
class Dispatcher(unohelper.Base, XJobExecutor):
    def __init__(self, ctx, *args):
        self.ComponentContext = ctx
        self.args = args
        fixPythonPath()

    def trigger(self, arg):
        try:
            reloadLeenoModules()
            ModFunc = arg.split('.')
            module = importlib.import_module(ModFunc[0])
            if module is None:
                print(f"Module '{ModFunc[0]}' not found")
                return

            func = getattr(module, ModFunc[1], None)
            if func is None:
                print(f"Function '{ModFunc[1]}' not found in Module '{ModFunc[0]}'")
                return

            if len(self.args) == 0:
                func()
            else:
                func(self.args)

        except Exception as e:
            handle_exception(e)

# Register the implementation
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Dispatcher,
    "org.giuseppe-vizziello.leeno.dispatcher",
    ("com.sun.star.task.Job",)
)
