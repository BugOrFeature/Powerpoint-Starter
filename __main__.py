import time
from os.path import isfile, join

import win32api
import win32gui
import os
import subprocess
import errno
import pyautogui

filesInDirectory = [f for f in os.listdir(os.getcwd()) if isfile(join(os.getcwd(), f))]
# filter in files with a pptx extension and exclude temporary files.
POWERPOINT_FILES = [f for f in filesInDirectory if '.pptx' in f and '~$' not in f]
START_IN_PRESENTATION_MODE = True
DEBUG = True

# change this path if its different on your pc
POWERPOINT_PATH = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.exe'


def start_presentation(file_name=None):
    try:
        # Open the process with the powerpoint path.
        subprocess.Popen([POWERPOINT_PATH, file_name])
    except OSError as e:
        if e.errno == errno.ENOENT:
            print("Powerpoint does not exist at:" + POWERPOINT_PATH)
        else:
            raise
            # Something else went wrong while trying to run the program\file


def maximiseWindowCallback(hwnd, args):
    """
    Will get called by win32gui.EnumWindows, once for each
    top level application window.
    """
    window_name = args[0]
    monitor = args[1]
    pyRect = monitor[2]
    try:
        # Get window title
        title = win32gui.GetWindowText(hwnd)
        if title.find(window_name) != -1:
            if DEBUG:
                print(f"Move window- x:{pyRect[0]}, y:{pyRect[2]}, width:{pyRect[3]} height:{pyRect[0]}")
            window = pyautogui.getWindowsWithTitle(window_name)
            curr = window.pop()
            # If the window is already maximized we have to restore it in order to avoid "double maximising" windows.
            if curr.isMaximized:
                curr.restore()
            width = abs(pyRect[2] - pyRect[0])
            if DEBUG:
                print("Move window")
            win32gui.MoveWindow(hwnd, pyRect[0], pyRect[1], width, pyRect[3], True)
            if DEBUG:
                print(f"Moved window- x:{pyRect[0]}, y:{pyRect[1]}, width:{width} height:{pyRect[3]}")
    except Exception as e:
        print(str(e))
        pass


def startPresentationCallback(hwnd, args):
    """
    Will get called by win32gui.EnumWindows, once for each
    top level application window.
    """
    window_name = args[0]
    print(f"Window name: {window_name}")
    print(f"Window text: {win32gui.GetWindowText(hwnd)}")
    print(f"Window matches window name: {win32gui.GetWindowText(hwnd).find(window_name)}")
    try:
        title = win32gui.GetWindowText(hwnd)
        if title.find(window_name) != -1:
            if DEBUG:
                print(f"Maximise and start presentation - window: {window_name}")
                print("start presentation callback")
            window = pyautogui.getWindowsWithTitle(window_name)
            time.sleep(.5)
            curr = window.pop()
            if not curr.isMaximized:
                curr.activate()
                curr.maximize()
            time.sleep(.05)
            pyautogui.keyDown('f5')
            time.sleep(.5)
            pyautogui.keyUp('f5')
            if DEBUG:
                print("maximise and start presentation")
    except Exception as e:
        print(str(e))
        pass


if __name__ == "__main__":
    cwd = os.getcwd()
    # Get monitor info
    monitors = []
    for hMonitor, hdcMonitor, pyRect in win32api.EnumDisplayMonitors():
        monitors.append((hMonitor, hdcMonitor, pyRect))

    if DEBUG:
        print(f"monitors: {monitors}")
        print(f"powerpoint files: {POWERPOINT_FILES}")
    # OPEN POWER POINTS
    for i in range(len(POWERPOINT_FILES)):
        time.sleep(.5)
        start_presentation(POWERPOINT_FILES[i])

    # MAXIMISE power points on monitors
    time.sleep(2)
    windows = pyautogui.getAllWindows()
    if DEBUG:
        print(f"windows: {windows}")

    powerpoint_name = [os.path.splitext(x)[0] for x in POWERPOINT_FILES]
    for i in range(len(monitors)):
        print("starting")
        win32gui.EnumWindows(maximiseWindowCallback, [powerpoint_name[i], monitors[i]])
        time.sleep(2)
        if START_IN_PRESENTATION_MODE:
            win32gui.EnumWindows(startPresentationCallback, [powerpoint_name[i], monitors[i]])
