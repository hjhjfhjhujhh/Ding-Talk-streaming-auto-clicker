
#from cv2 import imwrite
#import matplotlib.pyplot as plt
import numpy as np
import time
import os
import cv2
import win32gui
import win32api
#import win32con
from win32con import SW_SHOWNORMAL, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
import datetime

import sys
from PyQt5.QtWidgets import QApplication
import qimage2ndarray

print("===========================================")
print("钉钉直播点击器")
print("writen by HJ")
print("注意事项:\n钉钉窗口无法最小化，但是无需置于最顶层\n程序不能处理多群直播同时开的情况，暂时没有解决需求")
print("===========================================")

app = QApplication(sys.argv)  # QT模型初始化 用来处理截图
screen = QApplication.primaryScreen()


def get_resource_path(relative_path):  # 修正打包模板图片的临时路径
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


img_st_wt = get_resource_path("stream_white.jpg")
img_st_bl = get_resource_path("stream_black.jpg")
template_paths = [img_st_wt, img_st_bl]


def click(loc: list):
    '''鼠标点击 [x,y]'''
    #pyautogui.click(loc[0], loc[1])

    loc[0] = int(loc[0])
    loc[1] = int(loc[1])
    current_x, current_y = win32api.GetCursorPos()

    # 设置鼠标位置
    win32api.SetCursorPos((loc[0], loc[1]))

    # 点击鼠标左键
    win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, loc[0], loc[1], 0, 0)
    win32api.mouse_event(MOUSEEVENTF_LEFTUP, loc[0], loc[1], 0, 0)

    # 返回鼠标位置到原来的位置
    win32api.SetCursorPos((current_x, current_y))


def timefull() -> str:
    """返回[完整时间]"""
    return "[" + str(datetime.datetime.now()) + "]"


def screenshot(hwnd: int) -> np.ndarray:
    '''使用QT处理截图 传入要截图窗口的句柄 返回np.ndarray位图'''
    img = screen.grabWindow(hwnd).toImage()
    img = qimage2ndarray.rgb_view(img)
    return img


def find(image: np.ndarray, template: np.ndarray):
    '''在一张图片中寻找另一张图片。
    接收两个np.ndarray位图，返回坐标或者False'''
    zoom_factor_X = 0.5
    zoom_factor_Y = 0.5
    # 缩放图片(节省算力)
    image = cv2.resize(image, (0, 0), fx=zoom_factor_X, fy=zoom_factor_Y)
    template = cv2.resize(template, (0, 0), fx=zoom_factor_X, fy=zoom_factor_Y)

    # 创建模板匹配对象
    result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)

    # 定位最佳匹配
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # 置信度判断
    Threshold = 0.8
    # print(max_val)
    if max_val > Threshold:
        return [(max_loc[0] + template.shape[1]/2)/zoom_factor_X, (max_loc[1] + template.shape[0]/2)/zoom_factor_Y, max_val]
    else:
        return False


"""from cv2 import imwrite 
def find_diff_pixels(img1, img2):#这是查看图片变化用的测试函数
    # Ensure that the images have the same shape
    if img1.shape != img2.shape:
        raise ValueError("The images have different shapes")

    # Compare each pixel of the two images
    mask = np.not_equal(img1, img2)

    # Set all the pixels where the mask is True to red
    img1 = np.where(mask, [255,0,0], img1)

    return img1"""
##############################################################################################
NoChange = False  # 图像是否变化
img_last = None  # 上次循环截图
OnStreaming = False  # 是否已经在直播
templates = []
for path in template_paths:
    templates.append(cv2.imread(path)[:, :, ::-1])
    #imwrite("test.jpg", template)
print("初始化完成.")


while True:
    if NoChange:  # 图像无变化时加快捕捉频率
        time.sleep(0.1)
    else:
        time.sleep(1)

    hw = win32gui.FindWindow('StandardFrame_DingTalk', '钉钉')  # 获得钉钉窗口句柄
    if hw == 0:
        print(timefull()+"没有钉钉进程")
        continue

    win32gui.ShowWindow(hw, SW_SHOWNORMAL)  # 钉钉设置为前台
    lux, luy, rdx, rdy = win32gui.GetWindowRect(hw)  # 获得钉钉窗口位置，左上角与右下角的坐标
    img = screenshot(hw)  # 截图

    if np.array_equal(img_last, img):  # 检测图像是否变化，无变化则跳过
        if NoChange == False:
            # print(timefull()+"图像未变化...")
            NoChange = True
            #res = False
        continue
    else:
        print(timefull()+"捕捉到图像变动")
        #imwrite('different.jpg', find_diff_pixels(img, img_last))
        NoChange = False
        img_last = img
        # plt.imshow(img)
        # plt.show()
        for tem in templates:
            # plt.imshow(tem)
            # plt.show()
            res = find(img, tem)  # 寻找字符串
            if res != False:
                break

    if res != False:  # 查找到直播字符串
        if OnStreaming:
            continue  # 已经点击过直播按钮，跳过
        else:
            OnStreaming = True
            time.sleep(1)
            win32gui.SetForegroundWindow(hw)  # 置顶钉钉窗口
            click([lux+res[0], luy+res[1]])  # 点击字符串
            print(timefull()+"点击 " + str([lux+res[0], luy+res[1], res[2]]))
    else:
        OnStreaming = False
