#
# # -*- coding: utf-8 -*-
#

from ctypes import windll
user32 = windll.user32
user32.SetProcessDPIAware()
# 위의 세 줄로 스크린샷 좌표가 올바르게 나옴

from matplotlib import pyplot as plt
from PIL import ImageGrab
import win32gui

toplist, winlist = [], []

def enum_cb(hwnd, results):
    elem = (hwnd, win32gui.GetWindowText(hwnd))
    # print(elem[0], elem[1].encode('utf-8'))
    winlist.append(elem)

win32gui.EnumWindows(enum_cb, toplist)
keyword = '메모장'
applicationCandidate = [(hwnd, title) for hwnd, title in winlist if keyword in title.lower()]

# just grab the hwnd for first window matching applicationCandidate
for idx, each in enumerate(applicationCandidate):
    print(idx, each)

applicationCandidate = applicationCandidate[0]
hwnd = applicationCandidate[0]

win32gui.SetForegroundWindow(hwnd)

plt.ion()

while True:
    print(win32gui.GetWindowPlacement(hwnd), win32gui.GetWindowRect(hwnd), end='              v\r')
    bbox = win32gui.GetWindowRect(hwnd)
    img = ImageGrab.grab(bbox)

    # plt.imshow(img)
    # plt.show()
    plt.pause(.00001)

    import win32com.client as comclt
    wsh= comclt.Dispatch("WScript.Shell")
    wsh.AppActivate("Notepad") # select another application
    wsh.SendKeys("a") # send the keys you want
