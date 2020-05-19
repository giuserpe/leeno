import uno, unohelper, sys, types, importlib, builtins
from com.sun.star.task import XJobExecutor


#set this one to 0 for deploy mode
#leave to 1 if you want to disable python cache
#to be able to modify and run installed extension
DISABLE_CACHE = 1

#this class fakes XSCRIPTCONTEXT variable, as the services
#don't get it and we need it in old code
class MyScriptContext :
	
	def __init__(self, ctx):
		#CTX is XComponentContext - we store it
		self.ComponentContext = ctx
		
	def getComponentContext(self):
		return self.ComponentContext

	def getDesktop(self):
		return self.ComponentContext.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ComponentContext )
		
	def getDocument(self):
		return self.getDesktop().getCurrentComponent()

#menu dispatcher
class Dispatcher( unohelper.Base, XJobExecutor ):
	def __init__( self, ctx ):
		
		#CTX is XComponentContext - we store it
		self.ComponentContext = ctx
		
		#create a fake XSCRIPTCONTEXT
		self.ScriptContext = MyScriptContext(ctx)
		
		print("-----------------------------------------------")
		print("CREATED")
		print("-----------------------------------------------")

	def trigger( self, args ):

		#menu items are passed as module.function
		#so split them in 2 strings
		argv = args.split('.')
		
		#store XSCRIPTCONTEXT variable in global space
		#to be used by imported scripts
		if 'XSCRIPTCONTEXT' not in __builtins__:
			__builtins__ ['XSCRIPTCONTEXT'] = self.ScriptContext
			
		#locate the module from its name and check it
		module = importlib.import_module(argv[0])
		if module is None:
			print("Module '", argv[0], "' not found")
			return
		
		#reload the module if we don't want the cache
		if DISABLE_CACHE != 0:
			importlib.reload(module)
			
		module.XSCRIPTCONTEXT = self.ScriptContext

		#locate the function from its name and check it
		func = getattr(module, argv[1]);
		if(func is None):
			print("Function '", argv[1], "' not found in Module '", argv[0], "'")
			return

		#call the handler
		func(self.ComponentContext)

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
	Dispatcher,
	"org.giuseppe-vizziello.leeno.dispatcher",
	("com.sun.star.task.Job",),)
