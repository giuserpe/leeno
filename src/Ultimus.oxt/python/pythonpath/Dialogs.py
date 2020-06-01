'''
    Module to build dialogs in python
'''
import os
import inspect
from collections import namedtuple

import uno
import unohelper

from com.sun.star.awt import Rectangle
from com.sun.star.awt import Size
from com.sun.star.awt import XActionListener
from com.sun.star.task import XJobExecutor

from com.sun.star.awt import XUnitConversion
from com.sun.star.util import MeasureUnit

from LeenoConfig import Config
from LeenoUtils import getComponentContext, getDocument, createUnoService


def getCurrentPath():
    '''
    get current script's path
    '''
    return os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))


def getParentWindowInfo():
    '''
    Get point at desktop's center -- to be able to center dialogs around it
    '''
    ctx = uno.getComponentContext()
    oDesktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    oDoc = oDesktop.getCurrentComponent()

    oView = oDoc.getCurrentController()
    oWindow = oView.getFrame().getComponentWindow()
    rect = oWindow.getPosSize()
    Xc = int(rect.X + rect.Width / 2)
    Yc = int(rect.Y + rect.Height / 2)
    W = rect.Width
    H = rect.Height

    return Rectangle(Xc, Yc, W, H)


ScreenInfo = namedtuple('ScreenInfo', ['Width', 'Height', 'Display'])


def getScreenInfo():
    '''
    Get screen size
    '''
    ctx = uno.getComponentContext()
    # oDesktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    oToolkit = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.awt.Toolkit", ctx)
    aWorkArea = oToolkit.WorkArea
    nWidht = aWorkArea.Width
    nHeight = aWorkArea.Height
    # oWindow = oToolkit.getActiveTopWindow()
    # oDisplay = oWindow.Display
    # return ScreenInfo(nWidht, nHeight, oDisplay)
    return ScreenInfo(nWidht, nHeight, 0)


Scalef = namedtuple('Scalef', ['XScale', 'YScale'])


def getScaleFactors():
    '''
    Dialog positions are scaled by weird factors (2.625 and 2.25 on my machine)
    so we need to figure them out before proceeding
    
    '''
    doc = getDocument()
    docframe = doc.getCurrentController().getFrame()
    docwindow = docframe.getContainerWindow()

    sc = docwindow.convertSizeToPixel(Size(1000, 1000), MeasureUnit.APPFONT)

    return Scalef(1000.0 / float(sc.Width), 1000.0 / float(sc.Height))


def getBigIconSize():
    '''
    Get 'best' size for a big dialog icon
    (like the one of alert and ok dialogs)
    '''
    scInfo = getScreenInfo()
    siz = min(scInfo.Width, scInfo.Width)
    siz = int(siz / 20)
    return Size(siz, siz)


def getTextBox(txt):
    '''
    Get the size needed to display a multiline text box
    '''
    ctx = uno.getComponentContext()
    serviceManager = ctx.ServiceManager
    # toolkit = serviceManager.createInstanceWithContext("com.sun.star.awt.ExtToolkit", ctx)
    # dialogModel = serviceManager.createInstance("com.sun.star.awt.UnoControlDialogModel")
    # textModel = dialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")

    textModel = serviceManager.createInstance(
        "com.sun.star.awt.UnoControlFixedTextModel")
    text = serviceManager.createInstance(
        "com.sun.star.awt.UnoControlFixedText")
    text.setModel(textModel)
    text.setText(txt)
    size = text.getMinimumSize()
    textModel.dispose()
    return size


def getButtonSize(txt):
    '''
    Get 'best' button size in a dialog
    based on text
    '''
    size = getTextBox(txt)
    return Size(max(size.Width + 15, 100), size.Height + 15)


class BasicDialog(unohelper.Base, XActionListener, XJobExecutor):
    """
    Dialog Base Framework
    """

    def __init__(self, nPositionX=None, nPositionY=None, nWidth=None, nHeight=None, sTitle=None):

        unohelper.Base.__init__(self)

        self._LocalContext = uno.getComponentContext()
        self._ServiceManager = self._LocalContext.ServiceManager
        self._Toolkit = self._ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.ExtToolkit", self._LocalContext)

        # create dialog model and set its properties properties
        self._DialogModel = self._ServiceManager.createInstance(
            "com.sun.star.awt.UnoControlDialogModel")

        scales = getScaleFactors()
        self._DialogModel.PositionX = int(nPositionX * scales.XScale)
        self._DialogModel.PositionY = int(nPositionY * scales.YScale)
        self._DialogModel.Width = int(nWidth * scales.XScale)
        self._DialogModel.Height = int(nHeight * scales.YScale)

        self._DialogModel.Name = "Default"
        self._DialogModel.Closeable = True
        self._DialogModel.Moveable = True
        self._DialogModel.Title = sTitle
        self._DialogModel.DesktopAsParent = False

        # create the dialog container and set our dialog model into it
        self._DialogContainer = self._ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", self._LocalContext)
        self._DialogContainer.setModel(self._DialogModel)

        self._showing = False

    def addControl(self, sAwtName, sControlName, dProps):
        '''
        Add a control element to a dialog
        dProps are element's properties in a list, see classes below for examples
        '''
        # create the control
        oControlModel = self._DialogModel.createInstance(
            "com.sun.star.awt.UnoControl" + sAwtName + "Model")

        # still dont' know why, but we've got to scale items coordinates
        # so get the factors
        scales = getScaleFactors()

        # set control properties
        while dProps:
            prp = dProps.popitem()

            # scale positions / sizes if needed
            if prp[0] in ("PositionX", "Width"):
                uno.invoke(oControlModel, "setPropertyValue", (prp[0], int(prp[1] * scales.XScale)))
            elif prp[0] in ("PositionY", "Height"):
                uno.invoke(oControlModel, "setPropertyValue", (prp[0], int(prp[1] * scales.YScale)))
            else:
                uno.invoke(oControlModel, "setPropertyValue", (prp[0], prp[1]))
            oControlModel.Name = sControlName

        # insert the control into dialog
        self._DialogModel.insertByName(sControlName, oControlModel)

        # if the control is a Button, setup an event handler
        if sAwtName == "Button":
            self._DialogContainer.getControl(sControlName).addActionListener(self)
            self._DialogContainer.getControl(sControlName).setActionCommand(sControlName + '_OnClick')
        return oControlModel

    def runDialog(self):
        '''
        execute (runs) the dialog waiting for actions
        dialog must be closed by a control event
        '''
        # signal that we showed the dialog
        self._showing = True

        self._DialogContainer.setVisible(True)
        self._DialogContainer.createPeer(self._Toolkit, None)
        self._DialogContainer.execute()

    def showDialog(self):
        '''
        Display the dialog without waiting for action
        BEWARE, the dialog is NOT displayed instantly, but only
        after some time / operations on it
        '''

        # signal that we showed the dialog
        self._showing = True

        # self.DialogContainer.createPeer(self.Toolkit, None)
        self._DialogContainer.setVisible(True)

    def hideDialog(self):
        '''
        Hide the dialog displayed with showDialog()
        '''
        if not self._showing:
            return
        # signal that we showed the dialog
        self._showing = False

        self._DialogContainer.setVisible(False)

    def setOnTop(self):
        '''
        set dialog on top of others
        needed mostly for non-modal dialogs
        '''
        if not self._showing:
            return
        self._DialogContainer.toFront()
        self._DialogContainer.setFocus()

    def showing(self):
        '''
        check if dialog is showing
        '''
        return self._showing

    def closeable(self, c=True):
        '''
        make dialog (not) closeable
        '''
        self._DialogModel.Closeable = c

    def noCloseable(self):
        '''
        make dialog not closeable
        '''
        self.closeable(False)

    def moveable(self, c=True):
        '''
        make dialog (not) moveeable
        '''
        self._DialogModel.Moveable = c

    def noMoveable(self):
        '''
        make dialog not moveeable
        '''
        self.moveable(False)

    def sizeable(self, c=True):
        '''
        make dialog (not) sizeable
        '''
        self._DialogModel.Sizeable = c

    def noSizeable(self):
        '''
        make dialog not sizeable
        '''
        self.sizeable(False)

    def backgroundcolor(self, c):
        '''
        set dialog background color
        '''
        self._DialogModel.BackgroundColor = c


class ProgressBar(BasicDialog):
    '''
    A progress bar
    '''
    def __init__(self, title, message, cancelTitle=None, minVal=0, maxVal=100):

        # store the 'closeable' state
        self._closeable = (cancelTitle is not None)

        # try to get an "optimal" size from current window
        wi = getParentWindowInfo()
        Width = int(2 * wi.Width / 3)

        # correct the width just to be not too small nor too big
        screenInfo = getScreenInfo()
        if Width < screenInfo.Width / 8:
            Width = int(screenInfo.Width / 8)
        elif Width > screenInfo.Width / 4:
            Width = int(screenInfo.Width / 4)

        margin = 15

        messageSize = getTextBox(message + " (100.0%)")
        if Width < messageSize.Width + 2 * margin:
            Width = messageSize.Width + 2 * margin

        Height = messageSize.Height + 2 * margin

        # get some elements sizes
        if self._closeable:
            cancelSize = getButtonSize("Annulla")
            Height = Height + cancelSize.Height + margin
            if Width < cancelSize.Width + 2 * margin:
                Width = cancelSize.Width + 2 * margin
        else:
            cancelSize = Size(0, 0)

        progressSize = Size(Width - 2 * margin, 20)
        Height = Height + progressSize.Height + margin

        # we try to place the progress bar at center of parent window
        X = int(wi.X - Width / 2)
        Y = int(wi.Y - Height / 2)

        progressX = margin
        progressY = margin

        messageX = margin
        messageY = progressY + progressSize.Height + margin

        if self._closeable:
            cancelX = int((Width - cancelSize.Width) / 2)
            cancelY = messageY + messageSize.Height + margin

        BasicDialog.__init__(self, nPositionX=X, nPositionY=Y, nWidth=Width, nHeight=Height, sTitle=title)

        # store the message, we append the progress to it
        self._message = message
        self._minVal = minVal
        self._maxVal = maxVal

        dProgress = {
            "PositionX": progressX,
            "PositionY": progressY,
            "Width": progressSize.Width,
            "Height": progressSize.Height,
            "ProgressValueMin": minVal,
            "ProgressValueMax": maxVal, }
        self._progressBar = self.addControl("ProgressBar", "progressBar", dProgress)

        dMessage = {
            "PositionX": messageX,
            "PositionY": messageY,
            "Width": messageSize.Width,
            "Height": messageSize.Height,
            "Label": message,
            "Align": 0, }
        self._lbMessage = self.addControl("FixedText", "lbMessage", dMessage)

        if self._closeable:
            dCancel = {
                "PositionX": cancelX,
                "PositionY": cancelY,
                "Width": cancelSize.Width,
                "Height": cancelSize.Height,
                "Label": cancelTitle, }
            self._btnCancel = self.addControl("Button", "btnCancel", dCancel)

        self.moveable(False)
        self.sizeable(False)

        self._pos = minVal

    def actionPerformed(self, oActionEvent):
        '''
        event handler
        '''
        # if control is not closeable, just do nothing
        if not self._closeable:
            return

        # check if we pressed 'cancel' button
        if oActionEvent.ActionCommand == 'btnCancel_OnClick':
            # just hode the dialog
            self.hideDialog()

    def setProgress(self, pos):
        '''
        set progress bar value and update it on screen
        '''
        if not self._showing:
            return
        percent = '{:.0f}%'.format(100 * (pos - self._minVal) / (self._maxVal - self._minVal))
        txt = self._message + ' (' + percent + ')'
        self._lbMessage.Label = txt
        self._progressBar.ProgressValue = pos

        # just to be sure that the progress bar stays on top
        self.setOnTop()

        # store current position - we need it to change message or other stuffs
        self._pos = pos

    def setMessage(self, msg):
        '''
        sets/changes the message displayed along progress percent value
        '''
        self._message = msg
        self.setProgress(self._pos)

    def setLimits(self, minV, maxV, pos):
        '''
        sets/changes the min and max progressbar values
        and resets its position
        '''
        self._minVal = minV
        self._maxVal = maxV
        self._pos = pos
        self.setProgress(self._pos)


class MultiButton(BasicDialog):
    '''
    base class for multi-button dialogs
    'buttons' is a dictionary in the form of "name":retval
    '''
    def __init__(self, image, title, message, buttons):

        print("\n\nINIT\n\n\n")

        # compose the dialog by an image, a button and a text area
        margins = 15

        iconSize = getBigIconSize()

        buttonSize = Size(0, 0)
        for key, val in buttons.items():
            curButtonSize = getButtonSize(key)
            buttonSize = Size(max(buttonSize.Width, curButtonSize.Width), max(buttonSize.Height, curButtonSize.Height))

        allButtonsSize = Size((buttonSize.Width + margins) * len(buttons) - margins, buttonSize.Height)

        infoSize = getTextBox(message)

        Width = max(iconSize.Width + infoSize.Width + 3 * margins, allButtonsSize.Width + 2 * margins)
        Height = max(iconSize.Height, infoSize.Height) + buttonSize.Height + 3 * margins

        # we try to place the dialog bar at center of parent window
        wi = getParentWindowInfo()
        X = wi.X - int(Width / 2)
        Y = wi.Y - int(Height / 2)

        xIcon = margins
        yIcon = margins

        xInfo = xIcon + iconSize.Width + margins
        yInfo = margins

        if allButtonsSize.Width + 2 * margins == Width:
            xButton = margins
        else:
            xButton = int((Width - allButtonsSize.Width) / 2)
        yButton = Height - margins - buttonSize.Height

        BasicDialog.__init__(self, nPositionX=X, nPositionY=Y, nWidth=Width, nHeight=Height, sTitle=title)

        imgUrl = uno.systemPathToFileUrl(os.path.join(getCurrentPath(), image))
        dImage = {
            "PositionX": xIcon,
            "PositionY": yIcon,
            "Width": iconSize.Width,
            "Height": iconSize.Height,
            "ScaleImage": True,
            "ScaleMode": 1,
            "Border": 0,
            "ImageURL": imgUrl}
        self._lbImage = self.addControl("ImageControl", "_lbImage", dImage)

        dMessage = {
            "PositionX": xInfo,
            "PositionY": yInfo,
            "Width": infoSize.Width,
            "Height": infoSize.Height,
            "Label": message,
            "Align": 0, }
        self._lbMessage = self.addControl("FixedText", "lbMessage", dMessage)

        self._btns = []
        for key, val in buttons.items():
            dBtn = {
                "PositionX": xButton,
                "PositionY": yButton,
                "Width": buttonSize.Width,
                "Height": buttonSize.Height,
                "Label": key, }
            self._btns.append({'btn': self.addControl("Button", key, dBtn), 'value': val, 'name': key})
            xButton += buttonSize.Width + margins

        self.returnValue = None
        self.runDialog()

    def actionPerformed(self, oActionEvent):
        '''
        Close dialog when button pressed
        '''
        for btn in self._btns:
            if oActionEvent.ActionCommand == btn['name'] + '_OnClick':
                self._showing = False
                self.returnValue = btn['value']
                self._DialogContainer.endExecute()
                break


class Exclamation(MultiButton):
    '''
    Exclamation alert dialog with OK button
    '''
    def __init__(self, title, message):
        MultiButton.__init__(self, "exclamation.png", title, message, {'Ok': 0})


class Ok(MultiButton):
    '''
    Ok alert dialog with OK button
    '''
    def __init__(self, title, message):
        MultiButton.__init__(self, "ok.png", title, message, {'Ok': 0})


class Info(MultiButton):
    '''
    Ok alert dialog with OK button
    '''
    def __init__(self, title, message):
        MultiButton.__init__(self, "info.png", title, message, {'Ok': 0})

#####################################################################################


def FileSelect(titolo='Scegli il file...', est='*.*', mode=0):
    """
    titolo  { string }  : titolo del FilePicker
    est     { string }  : filtro di visualizzazione file
    mode    { integer } : modalità di gestione del file

    Apri file:  `mode in(0, 6, 7, 8, 9)`
    Salva file: `mode in(1, 2, 3, 4, 5, 10)`
    see:('''http://api.libreoffice.org/docs/idl/ref/
            namespacecom_1_1sun_1_1star_1_1ui_1_1
            dialogs_1_1TemplateDescription.html''' )
    see:('''http://stackoverflow.com/questions/30840736/
        libreoffice-how-to-create-a-file-dialog-via-python-macro''')
    """
    estensioni = {'*.*': 'Tutti i file(*.*)',
                  '*.odt': 'Writer(*.odt)',
                  '*.ods': 'Calc(*.ods)',
                  '*.odb': 'Base(*.odb)',
                  '*.odg': 'Draw(*.odg)',
                  '*.odp': 'Impress(*.odp)',
                  '*.odf': 'Math(*.odf)',
                  '*.xpwe': 'Primus(*.xpwe)',
                  '*.xml': 'XML(*.xml)',
                  '*.dat': 'dat(*.dat)', }
    oFilePicker = createUnoService("com.sun.star.ui.dialogs.FilePicker")
    oFilePicker.initialize((mode, ))
    oDoc = getDocument()

    # try to get path from current document, if any
    # if not, look into config to fetch last used one
    oPath = oDoc.getURL()
    if oPath == '':
        oPath = uno.systemPathToFileUrl(Config().read('Generale', 'ultimo_percorso'))
    else:
        oPath = os.path.dirname(oPath)
    oFilePicker.setDisplayDirectory(oPath)

    oFilePicker.Title = titolo
    app = estensioni.get(est)
    oFilePicker.appendFilter(app, est)
    if oFilePicker.execute():
        oPath = uno.fileUrlToSystemPath(oFilePicker.getFiles()[0])
        Config().write('Generale', 'ultimo_percorso', os.path.dirname(oPath))
        return oPath
    return None
