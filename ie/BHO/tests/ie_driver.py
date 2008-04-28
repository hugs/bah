import win32com.client
import win32gui
import win32process
import win32con
import win32api

import pythoncom
from threading import Thread
import time

DEBUG = False

# Internet Explorer's Class Id
CLSID = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'

class UnexpectedDialogBox(Exception):
    def __init__(self, value):
        self.value = value
    
    def __str__(self):
        return repr(self.value)

class MessageDispatcher(Thread):
    
    def __init__ (self, handle, message, should_wait=True):
        Thread.__init__(self)
        self.browser = None
        self.handle = handle
        self.message= message
        self.all = []
    
    def run(self):
        pythoncom.CoInitialize()
        self.update_browser_list()
        # Grab the correct IE, based on handle (HWND)
        self.browser = [browser for browser in self.all if browser.HWND == self.handle][0]
        self.waitUntilReady()
        return eval("self.browser." + self.message)
    
    def waitUntilReady(self):
        while self.browser.Busy:
            # this is a naive approach, but works in the very simple cases
            # should also add a wait for "readyState == 'complete'", too.
            time.sleep(.20)
    
    def update_browser_list(self):
        all = []     
        temp_all = win32com.client.Dispatch(CLSID)
        ie_count = 0
        # TODO add better waiting logic here
        while ie_count < 1:
            try:
                temp_all = win32com.client.Dispatch(CLSID)
                time.sleep(.2)
            except:
                print "error in temp_all = win32com.client.Dispatch(CLSID)"
                continue
                
            try:
                temp_count = temp_all.Count
                ie_count = temp_count
                if ie_count == 0:
                    print "ie_count still zero.. continuing..."
                    continue
                else:
                    print ie_count
                    break
            except:
                continue
        
        for i in range(ie_count):
            try:
                pw = temp_all[i].Document.parentWindow
                all.append(temp_all[i])
            except AttributeError:
                pass
        self.all = all


class Browser:
    def __init__(self):
        if DEBUG:
            # threadlist is useful for analyzing blocked threads during debugging..
            self.threadlist = []
        self._timeout = 5      # wait no more than 5 seconds before responding
        self._sleep_increment = .2  # sleep in 1/5 second increments
        
        # just a shortcut for .script
        self.js = self.script
        
        app = win32com.client.DispatchEx("InternetExplorer.Application")
        app.Navigate("about:blank")
        app.Visible = 1
        self.handle = app.HWND
        self.focus()
        self.should_wait = True        
        
    def alert(self, text_message):
        alert_message = "alert('" + text_message + "')"
        self.script(alert_message)
    
    def bho(self, bho_code):        
        bho_message = 'window.bho.' + bho_code
        self.script(bho_message)
    
    def focus(self):
        # this doesn't always work... :-(
        win32gui.SetForegroundWindow(self.handle)    
    
    def goto(self, url):
        message = 'Navigate("' + url + '")'
        self._dispatch_message(message)
    
    def quit(self):
        # The Old Way
        #   It didn't work if an alert box was present.
        #message = "Quit()"
        #self._dispatch_message(message)
        
        # The New Way
        #   Kills the browser -- no matter what!
        #   See http://www.velocityreviews.com/forums/t363843-kill-process-based-on-window-name-win32.html
        
        # Get the window's process id's
        t, p = win32process.GetWindowThreadProcessId(self.handle)
        
        # Ask window nicely to close
        win32gui.PostMessage(self.handle, win32con.WM_CLOSE, 0, 0)
        
        # Allow some time for app to close
        time.sleep(2)
        
        # If app didn't close, force close
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, p)
            if handle:
                win32api.TerminateProcess(handle,0)
                win32api.CloseHandle(handle)
        except:   # NAKED exception.. bad, bad, bad...
            pass
    
    def script(self, code):
        javascript_code = 'Document.parentWindow.execScript("' + code + '")'
        self._dispatch_message(javascript_code)
        

    
    # Private methods
    def _dispatch_message(self, message):
        new_message = MessageDispatcher(self.handle, message, self.should_wait)
        new_message.setDaemon(True)
        new_message.start()
        if DEBUG:
            self.threadlist.append(new_message)
        
        counter = 0
        i_should_keep_trying = True
        while i_should_keep_trying:         
            counter += 1   
            
            if new_message.isAlive() and (counter >= (self._timeout / self._sleep_increment)):
                #raise UnexpectedDialogBox("I'm freaking out! There's a dialog box!")
                # I'm freaking out! No test automation tool expects the alert box inquisition!
                print "Something is wrong. An unexpected dialog box, perhaps?"
                i_should_keep_trying = False
                break            
            
            if not new_message.isAlive():
                # the thread ended cleanly, no more need to wait...
                if DEBUG:
                    self.threadlist.remove( self.threadlist[-1] )
                i_should_keep_trying = False
                break
                          
            time.sleep( self._sleep_increment )                


if __name__ == '__main__':
    browser = Browser()
    browser.goto("http://localhost:8000/bho-tests.html")
    browser.script("document.getElementById('set-file-field').click()")  
    browser.script("document.getElementById('say-hello').click()")   
    time.sleep(5)
    browser.quit()


"""
doctests:

>>> from ie_driver import Browser
>>> browser = Browser()
>>> browser.alert("hello")
>>> browser.goto("http://localhost:8000/bho-tests.html")
>>> browser.quit()

"""

