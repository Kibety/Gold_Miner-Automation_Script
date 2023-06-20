# -*- coding: utf-8 -*-
# Copyright (C) 2023 # @version
# @Time    : 2023/4/12 14:55
# @Author  : Oraer
# @File    : toolkit.py
import time
import numpy as np
import win32gui, win32ui, win32con


class WindowCapture:
    # properties
    w = 0
    h = 0
    hwnd = None
    cropped_x = 0
    cropped_y = 0
    offset_x = 0
    offset_y = 0

    # constructor
    def __init__(self, window_name):
        # find the hwnd for the window we want to capture
        self.hwnd = win32gui.FindWindow(None, window_name)
        if not self.hwnd:
            raise Exception('Window not found: {}'.format(window_name))

        # 将该窗口最大化
        win32gui.ShowWindow(self.hwnd, win32con.SW_MAXIMIZE)
        # 设置窗口为最前
        # win32gui.SetForegroundWindow(self.hwnd)
        # 获取窗口的size
        window_rect = win32gui.GetWindowRect(self.hwnd)
        self.w = window_rect[2] - window_rect[0]
        self.h = window_rect[3] - window_rect[1]

        # 考虑窗口边框和标题栏并将它们剪掉
        border_pixels = 8
        titlebar_pixels = 30
        self.w = self.w - (border_pixels * 2)
        self.h = self.h - titlebar_pixels - border_pixels
        self.cropped_x = border_pixels
        self.cropped_y = titlebar_pixels

        # 设置裁剪的坐标偏移量，以便我们可以转换屏幕截图
        # 图像到实际的屏幕位置
        self.offset_x = window_rect[0] + self.cropped_x
        self.offset_y = window_rect[1] + self.cropped_y

    def get_screenshot(self):
        # get the window image data
        # 获取窗口图像数据
        # 根据窗口句柄获取窗口的设备上下文DC（Divice Context）
        wDC = win32gui.GetWindowDC(self.hwnd)
        # 根据窗口的DC获取dcObj
        dcObj = win32ui.CreateDCFromHandle(wDC)
        # dcObj创建可兼容的DC
        cDC = dcObj.CreateCompatibleDC()
        # 创建bigmap准备保存图片
        dataBitMap = win32ui.CreateBitmap()
        # 为bitmap开辟空间
        dataBitMap.CreateCompatibleBitmap(dcObj, self.w, self.h)
        # 高度saveDC，将截图保存到dataBitMap中
        cDC.SelectObject(dataBitMap)
        # 截取从左上角（0，0）长宽为（w，h）的图片
        cDC.BitBlt((0, 0), (self.w, self.h), dcObj, (self.cropped_x, self.cropped_y), win32con.SRCCOPY)

        # 将原始数据转换成opencv可以读取的格式
        # dataBitMap.SaveBitmapFile(cDC, 'debug.bmp')
        signedIntsArray = dataBitMap.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape = (self.h, self.w, 4)

        # 释放资源
        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())

        # drop the alpha channel, or cv.matchTemplate() will throw an error like:
        #   error: (-215:Assertion failed) (depth == CV_8U || depth == CV_32F) && type == _templ.type()
        #   && _img.dims() <= 2 in function 'cv::matchTemplate'
        img = img[..., :3]

        # make image C_CONTIGUOUS to avoid errors that look like:
        #   File ... in draw_rectangles
        #   TypeError: an integer is required (got type tuple)
        # see the discussion here:
        # https://github.com/opencv/opencv/issues/14866#issuecomment-580207109
        img = np.ascontiguousarray(img)

        return img

    def list_window_names(self):
        def winEnumHandler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                print(hex(hwnd), win32gui.GetWindowText(hwnd))
        win32gui.EnumWindows(winEnumHandler, None)

    def get_screen_position(self, pos):
        # translate a pixel position on a screenshot image to a pixel position on the screen.
        # 将屏幕截图图像上的像素位置转换为屏幕上的像素位置。
        # pos = (x, y)
        # WARNING: if you move the window being captured after execution is started, this will
        # return incorrect coordinates, because the window position is only calculated in
        # the __init__ constructor.
        return (pos[0] + self.offset_x, pos[1] + self.offset_y)


def get_window_hwnd(window_name):
    # 获取窗口句柄
    hwnd = win32gui.FindWindow(None, window_name)
    if not hwnd:
        raise Exception('Window not found: {}'.format(window_name))
    else:
        return win32gui.GetWindowRect(hwnd), hwnd  # 返回坐标值和handle


def get_current_window():
    """ 获取当前窗口
    :return:
    """
    return win32gui.GetForegroundWindow()


def set_current_window(hwnd):
    """ 强制当前窗口全屏
    :param hwnd:
    :return:
    """
    win32gui.SetForegroundWindow(hwnd)


def get_window_title(hwnd):
    """获取窗口的标题栏
    :param hwnd:
    :return:
    """
    return win32gui.GetWindowText(hwnd)


def get_current_window_title():
    """获取当前窗口的标题栏
    :return:
    """
    return get_window_title(get_current_window())


def find_window_by_title(title):
    """按照标题栏查找窗口
    :param title:
    :return: 所查找窗口的句柄
    """
    try:
        return win32gui.FindWindow(None, title)
    except Exception as ex:
        print('error calling win32gui.FindWindow ' + str(ex))
        return -1


def _get_all_hwnd(hwnd, mouse):
    """ 获取所有窗口的句柄及标题
    :param hwnd:
    :param mouse:
    :return:
    """
    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})


def get_all_hwnd(hwnd=None):
    global hwnd_title
    hwnd_title = dict()
    win32gui.EnumWindows(_get_all_hwnd, 0)
    for wnd in hwnd_title.items():
        print(wnd)


def get_window_pos(name):
    name = name
    hwnd = win32gui.FindWindow(0, name)
    # 获取窗口句柄
    if hwnd == 0:
        return None
    else:
        return win32gui.GetWindowRect(hwnd), hwnd     # 返回坐标值和handle


# if __name__ == "__main__":
    # 获取当前窗口句柄(是一个整数)
    # print(get_current_window())
    # # 获取当前窗口标题
    # print(get_current_window_title())
    #
    # 给定一个标题, 查找这个窗口, 如果找到就放到最前
    # hwnd = find_window_by_title('黄金矿工中文版小游戏,在线玩,4399小游戏')
    # set_current_window(hwnd)
    # time.sleep(0.5)
    # 打印刚刚切换到最前的窗口标题
    # print(get_current_window_title())
    #
    # 获取所有窗口的句柄及标题
    # hwnd_title = dict()
    # win32gui.EnumWindows(_get_all_hwnd, 0)
    # for wnd in hwnd_title.items():
    #     print(wnd)
    # 或者 get_all_hwnd()
