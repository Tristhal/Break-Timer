from win32api import *
from win32gui import *
import win32com.client
import win32gui as win32gui
import win32con
import sys, os
import struct
import time
import datetime
import ctypes
import re

class WindowsBalloonTip:
    "Base code from https://gist.github.com/wontoncc/1808234"
    def __init__(self):
        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
        }
        # Register the Window class.
        wc = WNDCLASS()
        self.hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbar"
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        self.classAtom = RegisterClass(wc)

    def ShowWindow(self, title, msg):
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( self.classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, self.hinst, None)
        UpdateWindow(self.hwnd)
        iconPathName = os.path.abspath(os.path.join( sys.path[0], "balloontip.ico" ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
           hicon = LoadImage(self.hinst, iconPathName, \
                    win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
          hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, "tooltip")
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, \
                         (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20,\
                          hicon, "Balloon  tooltip",msg,200,title))
        # self.show_balloon(title, msg)
        DestroyWindow(self.hwnd)

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0) # Terminate the app.

def setFront(windowID):
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.ShowWindow(windowID,5)
    win32gui.SetForegroundWindow(windowID)

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

def toFront(name):
    window = win32gui.GetForegroundWindow()
    top_windows = []
    win32gui.EnumWindows(windowEnumerationHandler, top_windows)
    for i in top_windows:
        if name in i[1].lower():
            # print(i)
            setFront(i[0])
            break
    return window

def strfdelta(tdelta, fmt):
    '''
    Takes a time.timedelta and returns a formatted string according to fmt. Format may contain {days}, {hours}, {minutes} and {seconds}
    '''
    d = {"days": tdelta.days}
    d["hours"], rem = divmod(tdelta.seconds, 3600)
    d["minutes"], d["seconds"] = divmod(rem, 60)
    return fmt.format(**d)

# Setting window title to be found in the search
window_title = "Break Timer"
os.system(f"title {window_title}")

# Instantiating Notifications
notification_baloon = WindowsBalloonTip()
minutes = int(input("How many minutes in between breakes: "))
lastInterval = time.time()

while True:
    print(strfdelta(datetime.timedelta(seconds=60 * minutes - abs(time.time() - lastInterval)), "Time Remaining: {hours} hours : {minutes} minutes : {seconds} seconds"))
    time.sleep(1)
    os.system("cls")
    if(abs(time.time() - lastInterval) >= 60 * minutes):
        last_window = toFront(window_title.lower())

        if last_window is None:
            print("Not __main__")

        notification_baloon.ShowWindow("Take a Break!", "Hit enter when you are done your break")

        ans = input("Hit enter when you are done your break or type exit to exit!")

        if re.search("\W*exit\W*", ans) is not None:
            setFront(last_window)
            break

        print("Continuing..")
        setFront(last_window)
        lastInterval = time.time()

