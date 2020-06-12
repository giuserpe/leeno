'''
Modulo di debug per LeenO
permette il debug attraverso l' IDE Eric6 (o simili)
'''
import sys
import os
import inspect
import subprocess
import time
import atexit

import inspect
import importlib

#
from os import listdir
from os.path import isfile, join

import unohelper
from com.sun.star.task import XJobExecutor
#

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

# openoffice path
# adapt to your system
_sofficePath = '/usr/lib/libreoffice/program'

OPENOFFICE_PORT = 8100
OPENOFFICE_PATH    = _sofficePath
OPENOFFICE_BIN     = os.path.join(OPENOFFICE_PATH, 'scalc')
OPENOFFICE_LIBPATH = OPENOFFICE_PATH

class OORunner:
    """
    Start, stop, and connect to OpenOffice.
    """
    def __init__(self, port=OPENOFFICE_PORT):
        """ Create OORunner that connects on the specified port. """
        self.port = port


    def connect(self, no_startup=False):
        """
        Connect to OpenOffice.
        If a connection cannot be established try to start OpenOffice.
        """
        localContext = uno.getComponentContext()
        resolver     = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        context      = None
        did_start    = False

        n = 0
        while n < 6:
            try:
                context = resolver.resolve("uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % self.port)
                break
            except NoConnectException:
                pass

            # If first connect failed then try starting OpenOffice.
            if n == 0:
                # Exit loop if startup not desired.
                if no_startup:
                     break
                self.startup()
                did_start = True

            # Pause and try again to connect
            time.sleep(1)
            n += 1

        if not context:
            raise Exception("Failed to connect to OpenOffice on port %d" % self.port)

        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

        if not desktop:
            raise Exception("Failed to create OpenOffice desktop on port %d" % self.port)

        if did_start:
            _started_desktops[self.port] = desktop

        return {'context': context,  'desktop':  desktop}


    def startup(self):
        """
        Start a headless instance of OpenOffice.
        """
        args = [OPENOFFICE_BIN,
                '--accept=socket,host=localhost,port=%d;urp;StarOffice.ServiceManager' % self.port,
                '--norestore',
                '--nofirststartwizard',
                '--nologo',
        #`        '--headless',
                ]
        env = os.environ.copy()
        # env  = {'PATH'       : '/bin:/usr/bin:%s' % OPENOFFICE_PATH, 'PYTHONPATH' : OPENOFFICE_LIBPATH, }

        try:
            # Open connection to server
            child = subprocess.Popen(args=args, env=env, start_new_session=False)
        except Exception as e:
            raise Exception("Failed to start OpenOffice on port %d: %s" % (self.port, e))

        #if pid <= 0:
        if child is None:
            raise Exception("Failed to start OpenOffice on port %d" % self.port)



    def shutdown(self):
        """
        Shutdown OpenOffice.
        """
        try:
            if _started_desktops.get(self.port):
                _started_desktops[self.port].terminate()
                del _started_desktops[self.port]
        except Exception:
            pass



# Keep track of started desktops and shut them down on exit.
_started_desktops = {}

def _shutdown_desktops():
    """ Shutdown all OpenOffice desktops that were started by the program. """
    for port, desktop in _started_desktops.items():
        try:
            if desktop:
                desktop.terminate()
        except Exception:
            pass

atexit.register(_shutdown_desktops)

# builtins dictionary in portable way... sigh
if type(__builtins__) == type(sys):
    bDict = __builtins__.__dict__
else:
    bDict = __builtins__


def reloadLeenoModules():
    '''
    This function reload all Leeno modules found in pythonpath
    '''
    # get our pythonpath
    myPath = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
    myPath = os.path.join(myPath, "pythonpath")

    # we need a listing of modules. We look at pythonpath ones
    pythonFiles = [f[: -3] for f in listdir(myPath) if isfile(join(myPath, f)) and f.endswith(".py")]
    for f in pythonFiles:
        print("Loading module:", f)
        module = importlib.import_module(f)

        # add to global dictionary, so it's available everywhere
        bDict[f] = module

        # reload the module
        importlib.reload(module)

###########################################################

# create the runner object
runner = OORunner()

# start libreoffice and get its context and desktop objects
lo = runner.connect()

# add our path and pythonpath subpath to our python path
leenoPath = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
sys.path.append(leenoPath)
leenoPath = os.path.join(leenoPath, "pythonpath")
sys.path.append(leenoPath)

# if we don't do so, we'll get a null current document
frames = lo['desktop'].getFrames()
if len(frames) > 0:
    frames[0]. activate()

'''
Poi sembra strano quando dico che il python è stato studiato con i piedi...

By default, when in the __main__ module, __builtins__ is the built-in module __builtin__ (note: no 's'); when in any other module,
__builtins__ is an alias for the dictionary of the __builtin__ module itself.
Note that in Python3, the module __builtin__ has been renamed to builtins to avoid some of this confusion.
'''

# setup our context for LeenO
bDict['__global_context__'] = lo['context']

# load LeenO modules
reloadLeenoModules()

import Dialogs

'''
Test dialog handler
'''
def testHandler(dialog,  widgetId,  widget,  cmdStr):
    print(f"Handler called with {cmdStr} sent by {widgetId}\n")
    w = dialog.getWidget("MyRadioGroup")
    n = w.getCount()
    i = w.getCurrent() + 1
    if i >= n:
        i = 0
    w.setCurrent(i)
    return False


'''
Test dialog
'''
'''
wr,  hr = Dialogs.getRadioButtonSize("X")

dlg = Dialogs.Dialog(Title='Ciao pepp',  Horz=False, CanClose=True,  Handler=testHandler,   Items=[
    Dialogs.HSizer(Items=[
        Dialogs.ImageControl(Image='info.png'),
        Dialogs.FixedText(Text="This is a nice, really really nice wanderful\nText Box\nwith some lines in it")
    ]),
    Dialogs.Spacer(),
    Dialogs.HSizer(Items=[
        Dialogs.RadioGroup(Id="MyRadioGroup", Default = 2,  Items=[
            "First radio button",
            "Second radio button",
            "Third radio button",
            "Another radio button"
        ]),
        Dialogs.Spacer(),
        Dialogs.VSizer(Items=[
            Dialogs.FixedText(Text="9999.99",  MinHeight=hr),
            Dialogs.FixedText(Text="9999.99",  MinHeight=hr),
            Dialogs.FixedText(Text="9999.99",  MinHeight=hr),
            Dialogs.FixedText(Text="9999.99",  MinHeight=hr),
        ]) ,
        Dialogs.Spacer(),
        Dialogs.RadioGroup(Id="MySecondRadioGroup", Items=['Some',  'More',  'Radio',  'Buttons']),
    ]),
    Dialogs.Spacer(),
    Dialogs.Spacer(),
    Dialogs.HSizer(Items=[
        Dialogs.Button(Label='Ok',  RetVal=1,  Icon='ok_24x24.png'),
        Dialogs.Spacer(),
        Dialogs.Button(Label='Cancel',  RetVal=0,  Icon='cancel_24x24.png'),
    ]),
])

dlg._layout()

res = dlg.run()

'''

'''
prg = Dialogs.Progress(Title="Progress test", Text="Sto lavorando", Closeable=True)

prg.show()
for i in range(1, 100):
    prg.setValue(i)
    if not prg.showing():
        print("CANCELLED")
        break
    time.sleep(0.05)
prg.hide()
'''

'''
a = True
radioH = Dialogs.getRadioButtonHeight()
imgW = Dialogs.getBigIconSize()[0] * 2
dlg = Dialogs.Dialog(Title="Importa dal formato XPWE", Items=[
    Dialogs.HSizer(Items=[
        Dialogs.VSizer(Items=[
            Dialogs.Spacer(),
           Dialogs.ImageControl(Id="img", Image="Icons-Big/question.png", MinWidth=imgW),
            Dialogs.Spacer()
        ]),
        Dialogs.VSizer(Id="vsizer", Items=[
            Dialogs.GroupBox(Label="Scegli elaborato", Items=[
                Dialogs.HSizer(Items=[
                    Dialogs.RadioGroup(Id="elab", Items=[
                        "Computo",
                        "Variante",
                        "Contabilità",
                        "Elenco Prezzi"
                    ]),
                    Dialogs.Spacer(),
                    Dialogs.VSizer(Items=[
                        Dialogs.FixedText(Id="TotComputo", Text="€ 999999.00", MinHeight=radioH),
                        Dialogs.FixedText(Id="TotVariante", Text="€ 999999.00", MinHeight=radioH),
                        Dialogs.FixedText(Id="TotContabilità", Text="€ 999999.00", MinHeight=radioH),
                        Dialogs.Spacer()
                    ])
                ])
            ])
        ])
    ])
])
if a:
    dlg["vsizer"].add(
        Dialogs.Spacer(),
        Dialogs.GroupBox(Label="Scegli destinazione", Items=[
            Dialogs.RadioGroup(Id="dest", Items=[
                "Documento corrente",
                "Nuovo documento"
            ])
        ])
    )

dlg["vsizer"].add(
    Dialogs.Spacer(),
    Dialogs.HSizer(Items=[
        Dialogs.Button(Label="Ok", RetVal=1, Icon="Icons-24x24/ok.png"),
        Dialogs.Spacer(),
        Dialogs.Button(Label="Annulla", RetVal=-1, Icon="Icons-24x24/cancel.png"),
    ])
)

dlg._layout()
x = dlg.dump()

dlg.run()
'''

from LeenoImport_XPWE import MENU_XPWE_import as X
X()

print("\nDONE\n")
