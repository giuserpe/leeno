import uno
 
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
 
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
 
"""https://wiki.openoffice.org/wiki/PythonDialogBox"""

def TestMessageBox():
  doc = XSCRIPTCONTEXT.getDocument()
  parentwin = doc.CurrentController.Frame.ContainerWindow
 
  s = "This a message"
  t = "Title of the box"
  res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)
 
  s = res
  t = "Titolo"

  #~ MessageBox(parentwin, s, t, "infobox")
 
# Show a message box with the UNO based toolkit
def MessageBox(ParentWin, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
 
  ctx = uno.getComponentContext()
  sm = ctx.ServiceManager
  sv = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx) 
  myBox = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)
  return myBox.execute()
 
#g_exportedScripts = TestMessageBox,
