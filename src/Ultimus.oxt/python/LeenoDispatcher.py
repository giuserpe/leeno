'''
    LeenO menu and basic function dispatcher
'''

import sys
import os
import inspect
import importlib

import unohelper
from com.sun.star.task import XJobExecutor


# set this one to 0 for deploy mode
# leave to 1 if you want to disable python cache
# to be able to modify and run installed extension
DISABLE_CACHE = 1


def fixPythonPath():
    '''
    This function should fix python path adding it to current sys path
    if not already there
    Useless here, just kept for reference
    '''

    print('''
    ########################################################################
    ######################## TRY THE TRICK ########################
    ########################################################################
    ''')

    # dirty trick to have pythonpath added if missing
    myPath = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
    myPath = os.path.join(myPath, "/pythonpath")
    if myPath not in sys.path:
        sys.path.append(myPath)

    print("sys.path:", sys.path)
    print('''
    ########################################################################
    ######################## pyleenp loading ########################
    ########################################################################
    ''')


class Dispatcher(unohelper.Base, XJobExecutor):
    '''
    LeenO menu and basic function dispatcher
    '''

    def __init__(self, ctx, *args):

        # CTX is XComponentContext - we store it
        self.ComponentContext = ctx

        # store any passed arg
        self.args = args
        
        # just in case...
        fixPythonPath()

    def trigger(self, arg):
        '''
        This function gets called when a menu item is selected
        or when a basic function calls PyScript()
        '''

        # menu items are passed as module.function
        # so split them in 2 strings
        ModFunc = arg.split('.')

        # locate the module from its name and check it
        module = importlib.import_module(ModFunc[0])
        if module is None:
            print("Module '", ModFunc[0], "' not found")
            return

        # reload the module if we don't want the cache
        if DISABLE_CACHE != 0:
            importlib.reload(module)

        # locate the function from its name and check it
        func = getattr(module, ModFunc[1])
        if func is None:
            print("Function '", ModFunc[1], "' not found in Module '", ModFunc[0], "'")
            return

        # call the handler, depending of number of arguments
        if len(self.args) == 0:
            func()
        else:
            func(self.args)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Dispatcher,
    "org.giuseppe-vizziello.leeno.dispatcher",
    ("com.sun.star.task.Job",),)
