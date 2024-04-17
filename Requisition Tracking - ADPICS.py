#copies and pastes the requisition to micro focus

from pyautogui import press, typewrite, hotkey
import time
import pygetwindow as gw
import os
import win32gui
import win32process
import win32api

#opens notepad and micro focus
os.startfile('C:\\Users\\AM00192\\OneDrive - TDCJ\\Fiscal Docs\\Other\\Coding Stuff\\Requisitions Pending PO.txt')
time.sleep(.25)
os.startfile('C:\\Users\\AM00192\\OneDrive - TDCJ\\Fiscal Docs\\Other\\Coding Stuff\\M Focus.txt')
time.sleep(.25)

win_1 = gw.getWindowsWithTitle('Requisitions Pending PO.txt')[0]
win_2 = gw.getWindowsWithTitle('M Focus.txt')[0]

win_1.activate()
win_1.restore()
hotkey ('ctrl', 'end')
press ('home')

fg_win = win32gui.GetForegroundWindow()
fg_thread, fg_process = win32process.GetWindowThreadProcessId(fg_win)
current_thread = win32api.GetCurrentThreadId()
win32process.AttachThreadInput(current_thread, fg_thread, True) #attaches thread to foreground window
try:
    end_cursor = win32gui.GetCaretPos()
    print (end_cursor)
finally:
    win32process.AttachThreadInput(current_thread, fg_thread, False) #detaches thread from foreground window
    
hotkey ('ctrl', 'home')

#loops copy/paste process
while True:
    hotkey ('shift', 'end')
    hotkey ('ctrl', 'c')
    time.sleep (.01)
    win_2.activate()
    win_2.restore()
    press ('home')
    hotkey ('ctrl', 'v')
    win_1.activate()
    win_1.restore()
    press ('home')
    press ('down')

    fg_win = win32gui.GetForegroundWindow()
    fg_thread, fg_process = win32process.GetWindowThreadProcessId(fg_win)
    current_thread = win32api.GetCurrentThreadId()
    win32process.AttachThreadInput(current_thread, fg_thread, True)
    try:
        active_cursor = win32gui.GetCaretPos()
        print (active_cursor)
    finally:
        win32process.AttachThreadInput(current_thread, fg_thread, False)
    if active_cursor == end_cursor:
        hotkey ('shift', 'end')
        hotkey ('ctrl', 'c')
        win_2.activate()
        win_2.restore()
        press ('home')
        hotkey ('ctrl', 'v')
        win_1.activate()
        win_1.restore()
        press ('home')
        break
