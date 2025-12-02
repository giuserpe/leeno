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

def handle_exception(e):
    try:
        pir = uno.getComponentContext().getValueByName(
            '/singletons/com.sun.star.deployment.PackageInformationProvider')
        expath_url = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
        expath = uno.fileUrlToSystemPath(expath_url)

        code_file = os.path.join(expath, 'leeno_version_code')
        version_line = ''
        if os.path.exists(code_file):
            with open(code_file, 'r', encoding='utf-8') as f:
                version_line = f.readline().strip()

        msg = (
            f"OS: {sys.platform} / LibreOffice-{loVersion()} / {version_line}\n\n"
            f"Errore: {str(e)}\n\n"
        )

        sysinfo = sys.exc_info()
        tb = sysinfo[2]
        if tb:
            tbInfo = traceback.extract_tb(tb)[-1]
            filen = os.path.basename(tbInfo.filename)
            msg += (
                f"File: '{filen}'\n"
                f"Line: '{tbInfo.lineno}'\n"
                f"Function: '{tbInfo.name}'\n"
            )

            # ==========================================================
            # 🔥 APERTURA AUTOMATICA DEL FILE IN VS CODE SULLA RIGA DELL’ERRORE
            # ==========================================================
            try:
                import subprocess

                full_path = tbInfo.filename
                line = tbInfo.lineno

                vscode_path = os.path.expanduser(r"~\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe")

                if os.path.exists(vscode_path):
                    subprocess.Popen([vscode_path, "-g", f"{full_path}:{line}"])
                else:
                    msg += "\n(VS Code non trovato nel percorso previsto)\n"

            except Exception as opener_err:
                msg += f"\n(Impossibile aprire VS Code: {opener_err})\n"
            # ==========================================================

        msg += "+--"*20 + "\nBACKTRACE:\n"
        for bk in traceback.extract_tb(tb):
            filen = os.path.basename(bk.filename)
            msg += f"File: {filen}, Line: {bk.lineno}, Function: {bk.name}\n"
        msg += "\n"

        log_dir = os.path.join(expath, "pythonpath")
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, "leeno_error.log")
        with open(log_path, "a", encoding="utf-8") as logfile:
            logfile.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
            logfile.write(msg)
            logfile.write("-"*60 + "\n\n")

        user = os.environ.get("USERNAME", "").lower()
        if "giuserpe" in user:
            msg += "+--"*20 + "\n" + traceback.format_exc()

        msgbox(Title="Errore interno", Message=msg)

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
