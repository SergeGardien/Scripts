"""Convert all swf files in this directory to a pdf file using Firefox.

Note that some parameters in this script need to be adjusted according to
user's printer setup, screen resolution, etc.

Documentation of pywinauto:
 * https://pywinauto.github.io/docs/contents.html
An example script using pywinauto:
 * https://github.com/vsajip/pywinauto/blob/master/examples/SaveFromFirefox.py
"""


import re
import os
import warnings
import webbrowser
from time import sleep
from functools import partial
import time

import pywinauto as pwa
from pywinauto.application import Application


def sendkey_escape(string):
    """Escape `+ ^ % ~ { } [ ] ( )` by putting them within curly braces.

    Refer to sendkeys' documentation for more info:
         * https://github.com/zvodd/sendkeys-py-si/blob/master/doc/SendKeys.txt
           (Could not open the original site: rutherfurd.net/python/sendkeys/ )
    """
    return re.sub(r'([+^%~{}\[\]()])', r'{\1}', string)


# Using 32-bit python on 64-bit machine? Will get the following warning a lot:
# "UserWarning: 64-bit application should be automated using 64-bit Python
# (you use 32-bit Python)"
# Limit this warnings to only show once.
# The following line does not work as expected. See
# github.com/pywinauto/pywinauto/issues/125
warnings.filterwarnings(
    'once', message=r'.*64-bit application should.*', category=UserWarning
)
# Assume Firefox is already open.
app = Application().connect(title_re=".*Firefox")
firefox = app.MozillaFireFox.GeckoFPSandboxChildWindow
filenames = os.listdir()
for filename in filenames:
    if not filename.endswith('.swf'):
        continue
    # pdfname = filename[:-3] + 'pdf'
    # if pdfname in filenames:
    #     # Already there!
    #     continue
    # Assume the default application to open swf files is Firefox.
    webbrowser.open(filename)
    firefox.SetFocus()
    firefox.Wait('exists ready', timeout=5)
    time.sleep(1)
    firefox.RightClickInput(coords=(200, 200))
    firefox.Wait('ready', timeout=10)
    # Click "print" from the rightclick menu.
    firefox.ClickInput(coords=(210, 320))
    pwa.timings.WaitUntilPasses(
        timeout=10,
        retry_interval=1,
        func=partial(app.Connect, title='Print'),
        exceptions=pwa.findwindows.WindowNotFoundError,
    )
    #app.Print.Wait('ready active', 5)
    # The printing process depends on the default printer being used.
    app.Print.OK.Click()
   	#app.Print.WaitNot('exists', timeout=5)
    # pwa.timings.WaitUntilPasses(
    #     timeout=20,
    #     retry_interval=1,
    #     func=partial(app.Connect, title='Save As'),
    #     exceptions=pwa.findwindows.WindowNotFoundError,
    # )
 #    # Be wary that some characters such as "%" don't work correctly in Save As
 #    # dialogs. This code does not handle such awkwardness of MS Windows.
 #    app.SaveAS.ComboBox.SetFocus().TypeKeys(
 #        sendkey_escape(os.getcwd() + '\\' + pdfname), with_spaces=True
 #    )
	# #app1 = Application().connect(path = r"c:\windows\splwow64.exe")
 #    #app1.SaveAS.Save.Click()
 #    firefox.Wait('exists ready', timeout=5)
 #    # Focuse is lost to flash (bugzilla: 78414). Use mouse to close the tab.
 #    firefox.ClickInput(coords=(418, 16), absolute=True)
 #    firefox.WaitNot("exists", timeout=5)