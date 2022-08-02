import webbrowser
from os.path import abspath, join, split, isfile
from sys import argv, executable
from time import time

from win32api import GetModuleHandle, LoadCursor
from win32con import CS_VREDRAW, CS_HREDRAW, IDC_ARROW, COLOR_WINDOW, WS_OVERLAPPED, WS_SYSMENU, LR_LOADFROMFILE, \
    LR_DEFAULTSIZE, IMAGE_ICON, IDI_APPLICATION, WM_USER, CW_USEDEFAULT
from win32gui import WNDCLASS, RegisterClass, CreateWindow, UpdateWindow, LoadImage, LoadIcon, NIF_ICON, NIF_MESSAGE, \
    NIF_TIP, Shell_NotifyIcon, NIM_ADD, NIM_MODIFY, NIF_INFO, PumpMessages, NIM_DELETE, PostQuitMessage
from winerror import ERROR_CLASS_ALREADY_EXISTS

PARAM_DESTROY = 1028
PARAM_CLICKED = 1029


class ToastNotifier:
    def __init__(self, title, msg, url=None):
        self.clicked = False
        self.url = url
        self.title = title
        self.msg = msg
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = 'InspectionNotifier' + str(time()).replace('.', '')
        wc.style = CS_VREDRAW | CS_HREDRAW
        wc.hCursor = LoadCursor(0, IDC_ARROW)
        wc.hbrBackground = COLOR_WINDOW
        wc.lpfnWndProc = self.wnd_proc

        try:
            RegisterClass(wc)
        except error as err_info:
            if err_info.winerror != ERROR_CLASS_ALREADY_EXISTS:
                raise

        style = WS_OVERLAPPED | WS_SYSMENU
        self.hwnd = CreateWindow(
            wc.lpszClassName,
            'Inspection Notifier',
            style,
            0,
            0,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            0,
            0,
            hinst,
            None,
        )
        UpdateWindow(self.hwnd)
        self._ShowToast(title, msg)

    def _ShowToast(self, title, msg):
        hinst = GetModuleHandle(None)
        iconPathName = abspath(
            join(split(argv[0])[0], 'favicon.ico')
        )
        if not isfile(iconPathName):
            iconPathName = abspath(
                join(split(executable)[0], 'DLLs', 'pyc.ico')
            )
        if not isfile(iconPathName):
            iconPathName = abspath(
                join(split(executable)[0], '..\\PC\\pyc.ico')
            )
        if isfile(iconPathName):
            icon_flags = LR_LOADFROMFILE | LR_DEFAULTSIZE
            hicon = LoadImage(
                hinst, iconPathName, IMAGE_ICON, 0, 0, icon_flags
            )
        else:
            print('未找到favicon.ico文件，将使用默认图标。')
            hicon = LoadIcon(0, IDI_APPLICATION)

        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, WM_USER + 20, hicon, 'Inspection')
        try:
            Shell_NotifyIcon(NIM_ADD, nid)
        except error:
            print('任务栏图标添加失败，文件管理器可能未运行')
        if self.url:
            msg += '\n>>> 点击以打开网页'
        Shell_NotifyIcon(NIM_MODIFY, (self.hwnd, 0, NIF_INFO,
                                      WM_USER + 20,
                                      hicon, "Balloon Tooltip", msg, 200,
                                      title))
        PumpMessages()

    def ReshowToast(self):
        self._ShowToast(self.title, self.msg)

    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if lparam == PARAM_CLICKED:
            if self.url:
                webbrowser.open(self.url)
                self.clicked = True
            self.on_destroy(hwnd, msg, wparam, lparam)
        elif lparam == PARAM_DESTROY:
            if self.url:
                self.clicked = False
            self.on_destroy(hwnd, msg, wparam, lparam)

    def is_clicked(self):
        return self.clicked

    def on_destroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0)
