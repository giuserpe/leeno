# -*- coding: utf-8 -*-

import uno
import unohelper

#~ from com.sun.star.awt import Point
from com.sun.star.frame import XDispatchProvider, XDispatch
from com.sun.star.lang import XInitialization
import traceback
import pyleeno

class ScriptContext(unohelper.Base):
    def __init__( self, ctx, doc, inv ):
        self.ctx = ctx
        self.doc = doc
        self.inv = inv

    def getDocument(self):
        if self.doc:
            return self.doc
        else:
            return self.getDesktop().getCurrentComponent()

    def getDesktop(self):
        return self.ctx.ServiceManager.createInstanceWithContext( "com.sun.star.frame.Desktop", self.ctx )

    def getComponentContext(self):
        return self.ctx

    def getInvocationContext(self):
        return self.inv


#~ Lmajor= 3 #'INCOMPATIBILITA'
#~ Lminor= 19 #'NUOVE FUNZIONALITA'
#~ Lsubv= "1" #'CORREZIONE BUGS

class LeenO(unohelper.Base, XInitialization, XDispatch, XDispatchProvider):
    def __init__(self, ctx):
        self.ctx = ctx
        self.psm = ctx.ServiceManager
        self.doc = None

    def create(self, service):
        return self.psm.createInstance(service)

    # XInitialization
    def initialize(self, args):
      if len(args) > 0:
         self.frame = args[0]
         self.doc = self.frame
         # print("\n\nframe: {}".format(self.frame))

    # XDispatchProvider
    def queryDispatch(self, url, framename, searchflags):
        # print("\n\nqueryDispatch: url = {}".format(url))
        if url.Protocol == "giuserpe:":
            return self
        return None
    def queryDispatches(self, requests):
        pass

    # XDispatch
    def dispatch(self, url, args):
        try:
            # print("\nURL = {}".format(url))
            # print("\nARGS = {}".format(args))

            function = url.Path
            getattr(self, function)()
        except:
            try:
                pyleeno.__dict__['XSCRIPTCONTEXT'] = ScriptContext(self.ctx, None, None)
                getattr(pyleeno, function)()
            except:
                traceback.print_exc()

    def addStatusListener(self, control, url):
        pass
    def removeStatusListener(self, control, url):
        pass
# https://forum.openoffice.org/it/forum/viewtopic.php?f=27&t=9470
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(LeenO,                           # UNO object class
                                         "org.giuseppe-vizziello.leeno.impl", # implemenation name
                                         ("org.giuseppe-vizziello.leeno.impl",),) # list of implemented services
