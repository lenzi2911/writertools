# Continuously updating word count
import unohelper, uno, os, time
from com.sun.star.i18n.WordType import WORD_COUNT
from com.sun.star.i18n import Boundary
from com.sun.star.lang import Locale
from com.sun.star.awt import XTopWindowListener

#socket = True
socket = False
localContext = uno.getComponentContext()

if socket:
    resolver = localContext.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', localContext)
    ctx = resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
else: ctx = localContext

smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

waittime = 5 # seconds
goal = 0

def printOut(txt):
    if socket: print txt
    else:
        model = desktop.getCurrentComponent()
        text = model.Text
        cursor = text.createTextCursorByRange(text.getEnd())
        text.insertString(cursor, txt + '\r', 0)

def hotCount(st):
    '''Counts the number of words in a string.

    ARGUMENTS:

    str st: count the number of words in this string

    RETURNS:

    int: the number of words in st'''
    startpos = long()
    nextwd = Boundary()
    lc = Locale()
    lc.Language = 'en'
    numwords = 1
    mystartpos = 1
    brk = smgr.createInstanceWithContext('com.sun.star.i18n.BreakIterator', ctx)
    nextwd = brk.nextWord(st, startpos, lc, WORD_COUNT)
    while nextwd.startPos != nextwd.endPos:
        numwords += 1
        nw = nextwd.startPos
        nextwd = brk.nextWord(st, nw, lc, WORD_COUNT)

    return numwords

def updateCount(wordCountModel, percentModel):
    '''Updates the GUI.

    Updates the word count and the percentage completed in the GUI. If some
    text of more than one word is selected (including in multiple selections by
    holding down the Ctrl/Cmd key), it updates the GUI based on the selection;
    if not, on the whole document.'''

    model = desktop.getCurrentComponent()
    try:
        if not model.supportsService('com.sun.star.text.TextDocument'):
            return
    except AttributeError: return

    sel = model.getCurrentSelection()
    try: selcount = sel.getCount()
    except AttributeError: return

    if selcount == 1 and sel.getByIndex(0).getString == '':
        selcount = 0

    selwords = 0
    for nsel in range(selcount):
        thisrange = sel.getByIndex(nsel)
        atext = thisrange.getString()
        selwords += hotCount(atext)

    if selwords > 1: wc = selwords
    else:
        try: wc = model.WordCount
        except AttributeError: return
    wordCountModel.Label = str(wc)

    if goal != 0:
        pc_text = '(%.2f percent)' % (100 * (wc / float(goal)))
        percentModel.Label = pc_text

# This is the user interface bit. It looks more or less like this:

###############################
# Word Count            _ o x #
###############################
#        _____                #
# 451 of |500| (90.20 percent)#
#        -----                #
###############################

# The boxed `500' is the text entry box.

class WindowClosingListener(unohelper.Base, XTopWindowListener):
    def __init__(self):
        global keepGoing
        
        keepGoing = True
    def windowClosing(self, e):
        global keepGoing
        
        keepGoing = False
        e.Source.setVisible(False)

def addControl(controlType, dlgModel, x, y, width, height, label, name = None):
    control = dlgModel.createInstance(controlType)
    control.PositionX = x
    control.PositionY = y
    control.Width = width
    control.Height = height
    if controlType == 'com.sun.star.awt.UnoControlFixedTextModel':
        control.Label = label
    elif controlType == 'com.sun.star.awt.UnoControlEditModel':
        control.Text = label

    if name:
        control.Name = name
        dlgModel.insertByName(name, control)
    else:
        control.Name = 'unnamed'
        dlgModel.insertByName('unnamed', control)

    return control

def loopTheLoop(goalModel, wordCountModel, percentModel):
    global goal

    while keepGoing:
        try: goal = int(goalModel.Text)
        except: goal = 0
        updateCount(wordCountModel, percentModel)
        time.sleep(waittime)

if not socket:
    import threading
    class UpdaterThread(threading.Thread):
        def __init__(self, goalModel, wordCountModel, percentModel):
            threading.Thread.__init__(self)

            self.goalModel = goalModel
            self.wordCountModel = wordCountModel
            self.percentModel = percentModel

        def run(self):
            loopTheLoop(self.goalModel, self.wordCountModel, self.percentModel)

def wordCount(arg = None):
    '''Displays a continuously updating word count.'''
    dialogModel = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialogModel', ctx)

    dialogModel.PositionX = 300
    dialogModel.PositionY = 20
    dialogModel.Width = 128 + 14
    dialogModel.Height = 16
    dialogModel.Title = 'Word Count'

    lblWc = addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 6, 2, 25, 14, '', 'lblWc')
    lblWc.Align = 2 # Align right
    addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 33, 2, 10, 14, 'of')
    addControl('com.sun.star.awt.UnoControlEditModel', dialogModel, 45, 2, 25, 14, '', 'txtGoal')

    addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 72, 2, 50, 14, '(percent)', 'lblPercent')
    addControl('com.sun.star.awt.UnoControlFixedTextModel', dialogModel, 124, 2, 12, 14, '', 'lblMinus')

    controlContainer = smgr.createInstanceWithContext('com.sun.star.awt.UnoControlDialog', ctx)
    controlContainer.setModel(dialogModel)

    controlContainer.addTopWindowListener(WindowClosingListener())
    controlContainer.setVisible(True)
    goalModel = controlContainer.getControl('txtGoal').getModel()
    wordCountModel = controlContainer.getControl('lblWc').getModel()
    percentModel = controlContainer.getControl('lblPercent').getModel()

    if socket:
        loopTheLoop(goalModel, wordCountModel, percentModel)
    else:
        uthread = UpdaterThread(goalModel, wordCountModel, percentModel)
        uthread.start()

keepGoing = True
if socket:
    wordCount()
else:
    g_exportedScripts = wordCount,

# Disclaimer and license from acb's macros
#
# Standard disclaimer and MIT licence .
# These macros are copyright (c) 2003-4 Andrew Brown.
# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation
# files (the 'Software'), to deal in the Software without
# restriction, including without limitation the rights to use, copy,
# modify, merge, publish, distribute, sublicense, and/or sell copies
# of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.
