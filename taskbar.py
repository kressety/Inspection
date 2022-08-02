from os.path import abspath, join, split, isfile
from sys import executable, argv
from threading import Thread
from time import time, localtime, strftime, sleep

from win32api import GetModuleHandle, LOWORD, LoadCursor
from win32con import LR_LOADFROMFILE, LR_DEFAULTSIZE, IMAGE_ICON, IDI_APPLICATION, WM_USER, WM_RBUTTONUP, MF_STRING, \
    TPM_LEFTALIGN, WM_NULL, WM_DESTROY, WM_COMMAND, CS_VREDRAW, CS_HREDRAW, IDC_ARROW, COLOR_WINDOW, WS_OVERLAPPED, \
    WS_SYSMENU, CW_USEDEFAULT, MF_SEPARATOR, MF_GRAYED
from win32gui import RegisterWindowMessage, LoadImage, LoadIcon, NIF_ICON, NIF_MESSAGE, NIF_TIP, Shell_NotifyIcon, \
    NIM_ADD, error, NIM_DELETE, PostQuitMessage, DestroyWindow, CreatePopupMenu, AppendMenu, GetCursorPos, \
    SetForegroundWindow, TrackPopupMenu, PostMessage, WNDCLASS, RegisterClass, CreateWindow, UpdateWindow, PumpMessages
from winerror import ERROR_CLASS_ALREADY_EXISTS

from request import TaskList
from toast import ToastNotifier


class TaskbarGUI:
    def __init__(self):
        self._time = time()
        self._Thread = None
        self._ResponseList = [[] for Task in range(len(TaskList))]
        self._Run = True
        msg_TaskbarRestart = RegisterWindowMessage('TaskbarCreated')
        message_map = {
            msg_TaskbarRestart: self.OnRestart,
            WM_DESTROY: self.OnDestroy,
            WM_COMMAND: self.OnCommand,
            WM_USER + 20: self.OnTaskbarNotify,
        }
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = 'InspectionTaskbar'
        wc.style = CS_VREDRAW | CS_HREDRAW
        wc.hCursor = LoadCursor(0, IDC_ARROW)
        wc.hbrBackground = COLOR_WINDOW
        wc.lpfnWndProc = message_map

        try:
            RegisterClass(wc)
        except error as err_info:
            if err_info.winerror != ERROR_CLASS_ALREADY_EXISTS:
                raise

        style = WS_OVERLAPPED | WS_SYSMENU
        self.hwnd = CreateWindow(
            wc.lpszClassName,
            'Inspection Taskbar',
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
        self._DoCreateIcons()
        self._StartNotice()
        Thread(target=self._AutoUpdate).start()
        PumpMessages()

    def _DoCreateIcons(self):
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

    def _Update(self):
        NothingToUpdate = True
        for TaskIndex in range(len(TaskList)):
            UpdateList = TaskList[TaskIndex]()
            for Index in range(len(self._ResponseList[TaskIndex])):
                if not self._ResponseList[TaskIndex][Index].is_clicked():
                    self._ResponseList[TaskIndex][Index].ReshowToast()
            if type(UpdateList) == str:
                ToastNotifier('Inspection', UpdateList)
            else:
                if len(UpdateList) != 0:
                    NothingToUpdate = False
                    for UpdateItem in UpdateList:
                        Response = ToastNotifier(UpdateItem[0], UpdateItem[1], UpdateItem[2])
                        if not Response.is_clicked():
                            self._ResponseList[TaskIndex].append(Response)
        if NothingToUpdate:
            ToastNotifier('Inspection', '无更新\n下次更新时间: {}'.format(
                strftime('%Y-%m-%d %H:%M:%S', localtime(self._time + 3600))))
        else:
            ToastNotifier('Inspection', '更新完成\n下次更新时间: {}'.format(
                strftime('%Y-%m-%d %H:%M:%S', localtime(self._time + 3600))))

    def _StartNotice(self):
        self._time = time()
        ToastNotifier('Inspection', '已启动，即将进行更新\n更新时间: {}\n下次更新时间: {}'.format(
            strftime('%Y-%m-%d %H:%M:%S', localtime(self._time)),
            strftime('%Y-%m-%d %H:%M:%S', localtime(self._time + 3600))))
        self._Update()

    def _AutoUpdate(self):
        while self._Run:
            while localtime() != localtime(self._time + 3600):
                sleep(0.5)
                if not self._Run:
                    return
            self._time = time()
            self._Update()

    def OnRestart(self, hwnd, msg, wparam, lparam):
        self._DoCreateIcons()

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        self._Run = False
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0)

    def OnTaskbarNotify(self, hwnd, msg, wparam, lparam):
        if lparam == WM_RBUTTONUP:
            menu = CreatePopupMenu()
            AppendMenu(menu, MF_GRAYED, 0, '上次更新时间：')
            AppendMenu(menu, MF_GRAYED, 0, strftime('%Y-%m-%d %H:%M:%S', localtime(self._time)))
            AppendMenu(menu, MF_SEPARATOR, 0, '')
            AppendMenu(menu, MF_STRING, 1024, '手动更新')
            AppendMenu(menu, MF_STRING, 1025, '退出')
            pos = GetCursorPos()
            SetForegroundWindow(self.hwnd)
            TrackPopupMenu(
                menu, TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None
            )
            PostMessage(self.hwnd, WM_NULL, 0, 0)

        return 1

    def OnCommand(self, hwnd, msg, wparam, lparam):
        ID = LOWORD(wparam)
        if ID == 1024:
            self._time = time()
            self._Update()
        elif ID == 1025:
            DestroyWindow(self.hwnd)
        else:
            print('未知的指令：', ID)


def RunTaskbar():
    Thread(target=TaskbarGUI).start()
