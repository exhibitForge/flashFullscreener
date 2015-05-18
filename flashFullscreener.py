import sys
if sys.version_info <= (3, 0):
    print("Sorry, requires Python 3.x, not Python", str(sys.version_info.major) + "." + str(sys.version_info.minor), "\n")
    sys.exit(1)

import os
import socket
import time
import logging

logPath = 'C:/exf/logs'
os.makedirs(logPath, exist_ok=True)
logging.basicConfig(
    filename = os.path.join(logPath, 'flashFullscreenerService.log'),
    level = logging.DEBUG,
    format = '[flashFullscreenerService] %(asctime)s %(levelname)-7.7s %(message)s'
)

try:
    import win32serviceutil
    import win32service
    import win32event
    import servicemanager
    import win32gui
    import win32con
    import win32api
except ImportError as ie:
    logging.critical('Cannot import from pywin32: ', ie)
    logging.critical('  Try installing from http://sourceforge.net/projects/pywin32/files/pywin32/')
    exit()

#from flashFullscreener import FlashFullscreener

class FlashFullscreenerService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'flashFullscreenerService'
    _svc_display_name_ = 'flashFullscreenerService'
    _svc_description_ = 'makes sure flash is fullscreen'
    #_flashFullscreener = None

    def __init__(self,args):
        logging.info('** Service Installing **')
        win32serviceutil.ServiceFramework.__init__(self,args)
        self.stop_event = win32event.CreateEvent(None,0,0,None)
        socket.setdefaulttimeout(60)
        self.stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        logging.info('** Service Stopping **')
        self.stop_requested = True

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_,'')
        )
        self.main()

    def main(self):
        logging.info('** Service Starting **')
        self._flashFullscreener = FlashFullscreener()
        ffs = self._flashFullscreener
        count = 0
        while not self.stop_requested:
            logging.debug('count=%s' % count)
            ffs.findAndMaximizeThenFullscreen()
            count += 1
            time.sleep(5)
        logging.info('Stop signal was received.')
        return

class FlashFullscreener():

    def _windowEnumerationHandler(self, hwnd, top_windows):
        top_windows[win32gui.GetWindowText(hwnd)] = hwnd

    def findAndMaximizeThenFullscreen(self, windowString='Adobe Flash Player 9'):
        found = False
        topWindowsDict = {}
        win32gui.EnumWindows(self._windowEnumerationHandler, topWindowsDict)
        for title, hwnd in topWindowsDict.items():
            print('Found a window called "%s"' % title)
            if title==windowString: #this is the window your looking for ...
                print('#this is the window we\'re looking for ...')
                found = True
                logging.debug('Checking window "%s"' % windowString)
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE ) #see https://github.com/facelessuser/Pywin32/blob/master/lib/x32/win32/lib/win32con.py for SW options
                win32gui.SetForegroundWindow(hwnd)
                if(win32gui.GetWindowRect(hwnd)[0]!=0): #the left side of the window isn't zero => not fullscreen
                    logging.info('  window not fullscreen - toggling')
                    # inject 'Ctrl-F' into the window (flash) to fullscreen it
                    win32api.keybd_event(win32con.VK_LCONTROL,0x1d, 0, 0) # see http://stackoverflow.com/questions/9255030/using-python-and-pywin32-to-automate-data-entry
                    win32api.keybd_event(win32api.VkKeyScan('F'),0x1e, 0, 0)
                    win32api.keybd_event(win32api.VkKeyScan('F'),0x9e, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_LCONTROL,0x9d, win32con.KEYEVENTF_KEYUP, 0)
                return
            else:
                pass
        if not found:
            print('Can\'t find it')
            logging.warning('Cannot find a window titled "%s"' % windowString)

if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(FlashFullscreenerService)
