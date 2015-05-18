import sys
if sys.version_info <= (3, 0):
    print("Sorry, requires Python 3.x, not Python", str(sys.version_info.major) + "." + str(sys.version_info.minor), "\n")
    sys.exit(1)

import time

try:
    import win32gui
    import win32con
    import win32api
except ImportError as ie:
    print('Cannot import from pywin32: ', ie)
    print('  Try installing from http://sourceforge.net/projects/pywin32/files/pywin32/')
    exit()


def windowEnumerationHandler(hwnd, top_windows):
    top_windows[win32gui.GetWindowText(hwnd)] = hwnd

def findAndMaximizeThenFullscreen(windowString):
    results = []
    topWindowsDict = {}
    win32gui.EnumWindows(windowEnumerationHandler, topWindowsDict)
    for title, hwnd in topWindowsDict.items():
        if title==windowString: #this is the window your looking for ...
            print('Checking window "%s"' % windowString)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE ) #see https://github.com/facelessuser/Pywin32/blob/master/lib/x32/win32/lib/win32con.py for SW options
            win32gui.SetForegroundWindow(hwnd)
            if(win32gui.GetWindowRect(hwnd)[0]!=0): #the left side of the window isn't zero => not fullscreen
                print('  window not fullscreen - toggling')
                # inject 'Ctrl-F' into the window (flash) to fullscreen it
                win32api.keybd_event(win32con.VK_LCONTROL,0x1d, 0, 0) # see http://stackoverflow.com/questions/9255030/using-python-and-pywin32-to-automate-data-entry
                win32api.keybd_event(win32api.VkKeyScan('F'),0x1e, 0, 0)
                win32api.keybd_event(win32api.VkKeyScan('F'),0x9e, win32con.KEYEVENTF_KEYUP, 0)
                win32api.keybd_event(win32con.VK_LCONTROL,0x9d, win32con.KEYEVENTF_KEYUP, 0)
            break

if __name__ == "__main__":
    for i in range(5): #TODO
        findAndMaximizeThenFullscreen('Adobe Flash Player 9')
        time.sleep(1.0)
