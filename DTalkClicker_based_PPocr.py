import numpy as np
from paddleocr import PaddleOCR
import os
import time
#import pyautogui
from cv2 import imwrite
import win32gui
import win32api
import win32con
#import win32ui
import datetime
import sys
from PyQt5.QtWidgets import QApplication
import qimage2ndarray
from PIL import ImageGrab

JUMPWAIT = False  # 调试 跳过等待
DEBUG_OCROT = False  # 调试 输出OCR结果
DEBUG_SAVEIMG = False  # 调试 保存截图
DEBUG = False  # 调试 其他


def timefull():
    """返回[完整时间]"""
    return "[" + str(datetime.datetime.now()) + "]"


def timeshort():
    """返回[日期]"""
    return "[" + str(datetime.datetime.now().date()) + "]"


def console(command: str) -> str:
    """执行cmd命令并返回结果"""
    try:
        return os.popen(command).read().strip()
    except:
        r = os.popen(command)
        r = r.buffer.read().decode("utf-8").strip()
        return r


def click(loc: list):
    '''鼠标点击 [x,y]'''
    #pyautogui.click(loc[0], loc[1])

    loc[0] = int(loc[0])
    loc[1] = int(loc[1])
    current_x, current_y = win32api.GetCursorPos()

    # 设置鼠标位置
    win32api.SetCursorPos((loc[0], loc[1]))

    # 点击鼠标左键
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, loc[0], loc[1], 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, loc[0], loc[1], 0, 0)

    # 返回鼠标位置到原来的位置
    win32api.SetCursorPos((current_x, current_y))


class Log:
    def __init__(self, filename):
        """
        filename: 日志文件名
        """
        self.filename = filename
        self.file = open(self.filename, "a", encoding='utf-8')

    def isopen(self):
        return not self.file.closed

    def open(self):
        """打开日志文件"""
        self.file = open(self.filename, "a")

    def write(self, info):
        """写入日志"""
        closed = self.file.closed
        if closed:
            self.open()
        self.file.write(info)
        self.file.flush()
        if closed:
            self.close()
        return info

    def close(self):
        """关闭日志文件"""
        self.file.close()

    def __del__(self):
        """析构函数"""
        self.close()


def logprint(string, *logs, Print=False):
    """
    向日志文件输出
    :params:
        ``string``: 日志条目
        ``logs``: 输出的log对象
        ``logprint``: 是否打印到控制台
    """
    for log in logs:
        log.write(timefull()+string+"\n")
    if Print or DEBUG:
        print(timefull()+string)  # [debug]


def errlogprint(string, *logs, close=True):
    """
    向日志文件输出错误
    :params:
        ``string``: 日志条目
        ``logs``: 输出的log对象
        ``close``: 是否退出程序
    """
    for log in logs:
        log.write(timefull()+"[ERR]"+string+"\n\n\n")
        del log

    logprint(timefull()+string)  # [debug]
    if close:
        input("按回车键退出...")
        sys.exit()


class OCR_result:
    """一个OCR对象"""

    def __init__(self, OCR: list = "", callback=None, callbackERR=logprint):
        self.OCR: list = OCR
        """OCR结果 正常初始化为空 可以手动传入(用于继承)"""
        if callback == None:
            self.callback = lambda x: x
        else:
            self.callback = callback
        """日志输出函数 可以为空 这个class是我给另一个项目写的
        所以保留了这个功能 实际上这个脚本我只给它传入了logprint"""
        if callbackERR == None:
            self.callbackERR = lambda x: x
        else:
            self.callbackERR = callbackERR
        """错误日志函数"""

        self.img = None
        """用来存储截图图像的函数"""

    def get_img(self, LupX=0, LupY=0, RdnX=1920, RdnY=1080, rectify=-10):
        """获取截图(坐标)
        :params:
            LupX:左上角的X坐标
            LupY:左上角的Y坐标
            RdnX:右下角的X坐标
            RdnY:右下角的Y坐标
            rectify:修正参数，截图框向外扩大几个像素
        """
        LupX -= rectify
        LupY -= rectify
        RdnX += rectify
        RdnY += rectify
        # self.callback("截图获取.")

        """#这个注释里是使用pyautogui截图的方法，这个方法不支持负数坐标
        img = pyautogui.screenshot(region=(LupX, LupY, RupX, RupY))
        #保存截图：
        #img.save('screenshot1.jpg')
        img = np.array(img)[:, :, ::-1] 
        #方括号的作用是进行一个神奇的通道转换 把BRG格式转换成RGB
        #如果不这么干图片会变黄 不过不怎么影响OCR"""

        """# 使用win32库截图 因为会随机出问题已弃用
        for i in range(10):  # 这个循环是尝试解决BitBlt faild问题 多重试几次
            hwnd = win32gui.GetDesktopWindow()
            hwindc = win32gui.GetWindowDC(hwnd)
            srcdc = win32ui.CreateDCFromHandle(hwindc)
            memdc = srcdc.CreateCompatibleDC()
            bitmap = win32ui.CreateBitmap()
            bitmap.CreateCompatibleBitmap(srcdc, RdnX-LupX, RdnY-LupY)
            memdc.SelectObject(bitmap)

            try:
                memdc.BitBlt((0, 0), (RdnX-LupX, RdnY-LupY), srcdc, (LupX, LupY), win32con.SRCCOPY)
            except:
                pass
                #memdc.BitBlt((0, 0), (RdnX-LupX, RdnY-LupY), srcdc, (LupX, LupY), win32con.SRCCOPY)#《《《《《
                srcdc.DeleteDC()
                memdc.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwindc)
                win32gui.DeleteObject(bitmap.GetHandle())
                del hwnd, hwindc, srcdc, memdc, bitmap
                
                return False
            break

        # 转换为np.ndarray对象
        signedIntsArray = bitmap.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape = (RdnY-LupY, RdnX-LupX, 4)

        # 释放内存
        srcdc.DeleteDC()
        memdc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwindc)
        win32gui.DeleteObject(bitmap.GetHandle())
        del hwnd, hwindc, srcdc, memdc, bitmap, signedIntsArray"""

        # 使用PIL截图 可以处理多屏幕的情况 但是不能锁屏，不能远程控制
        img = ImageGrab.grab(bbox=(LupX, LupY, RdnX, RdnY), all_screens=True)
        img = np.array(img.getdata(), np.uint8).reshape(img.size[1], img.size[0], 3)
        img = np.array(img)[:, :, ::-1]  # 方括号的作用是进行一个神奇的通道转换 把BRG格式转换成RGB

        # 保存
        if DEBUG_SAVEIMG:
            imwrite('screenshot.jpg', img)
            #imwrite("screenshot.jpg", img[:, :, :3])
            #bitmap.SaveBitmapFile(memdc, "screenshot.jpg")
        self.img = img

        return self
    def get_img_fwindow(self,hwnd):
        """获取截图(窗口)
        :params:
            hwnd:窗口的句柄
        """
        
        app = QApplication(sys.argv)#使用qt截图 这一行不能删
        screen = QApplication.primaryScreen()
        img = screen.grabWindow(hwnd).toImage()
        img = qimage2ndarray.rgb_view(img)[:, :, ::-1]


        if DEBUG_SAVEIMG:
            imwrite('screenshot.jpg', img)
        self.img = img

        return self
    def ocr(self):
        """对类内保存的截图进行OCR"""
        self.callback("OCR识别.")
        self.OCR = ocr.ocr(self.img, cls=False)[0]
        return self

    def get(self, LupX=0, LupY=0, RdnX=1920, RdnY=1080):
        """
        获取截图,并进行OCR
        return: 自身
        """
        self.callback("OCR获取.")
        ret = self.get_img(LupX, LupY, RdnX, RdnY)
        if ret:
            self.OCR = ocr.ocr(self.img, cls=False)[0]
        return ret

    def setimage(self, img):
        self.OCR = ocr.ocr(img, cls=False)[0]
        return self

    def find(self, OCR: str, strong=False, retry=0) -> list:
        """
        定位OCR结果中某个字符串
        ''OCR'':要查找的文本
        ''strong'':强匹配，整个文本框完全匹配，不能多字
        ''retry'':重试次数
            return:
            ``place``: 字符串坐标与内容: [x:int , y:int  ,str:str]
            ``Flase``: 没找到
        """
        for i in range(retry + 1):
            for a in self.OCR:
                
                if OCR in a[1][0]:
                    if strong and a[1][0] != OCR:
                        self.callback("OCR定位: "+OCR+" >存在，但由于强判断舍弃")
                        continue
                    place = [0, 0, ""]
                    place[0] = (a[0][0][0]+a[0][2][0])/2
                    place[1] = (a[0][0][1]+a[0][2][1])/2
                    place[2] = a[1][0]
                    self.callback("OCR定位: "+OCR+" >存在:"+str(place))
                    return place
            self.callback("OCR定位: "+OCR+" >不存在")
            if i < retry:
                self.get()
        return False

    def have(self, OCR: str, strong=False, retry=0) -> bool:
        """
        判断OCR结果是否有某个字符串
        """
        for i in range(retry + 1):
            for a in self.OCR:
                if OCR in a[1][0]:
                    if strong and a[1][0] != OCR:
                        self.callback("OCR判断: "+OCR+" >存在，但由于强判断舍弃")
                        continue
                    self.callback("OCR判断: "+OCR+" >存在")
                    return True
            if i < retry:
                self.get()
                time.sleep(1)
            self.callback("OCR判断: "+OCR+" >不存在")
        return False

    def __str__(self) -> str:
        return str(self.OCR)


if DEBUG:  # 调试模式输出
    OCRR = OCR_result(callback=lambda x: logprint(x, mainlog), callbackERR=lambda x: errlogprint(x, mainlog))
else:
    OCRR = OCR_result(callback=None, callbackERR=lambda x: logprint(x, mainlog))

"""BTfailed = False  # <<<<<<<<<<<<<<<<<<<
FailedCount = 0"""
ocr = PaddleOCR(lang="ch", use_angle_cls=False, show_log=DEBUG)  # paddleocr初始化
mainlog = Log("主日志.log")  # 日志初始化
##############################################################################################
logprint("", mainlog)
logprint("初始化完成.", mainlog)
if not DEBUG:
    print("初始化完成...\n调试输出已关闭.")
jumpWait = True  # 是否跳过等待
img_last = None  # 储存上次的截图
res = res_last = [0, 0, ""]  # 储存目标字符串位置
while True:
    if (not jumpWait) and (not JUMPWAIT):  # 是否等待
        time.sleep(60)
    else:
        time.sleep(1)

    hw = win32gui.FindWindow('StandardFrame_DingTalk', '钉钉')  # 获得钉钉窗口句柄
    if hw == 0:
        logprint("没有钉钉进程", mainlog)
        continue
    elif win32gui.IsIconic(hw):  # 如果窗口最小化则设置为前台
        win32gui.ShowWindow(hw, win32con.SW_SHOWNORMAL)

    lux, luy, rdx, rdy = win32gui.GetWindowRect(hw)  # 获得钉钉窗口位置，左上角与右下角的坐标
    """
    if BTfailed == True:

        if OCRR.get_img(lux, luy, rdx, rdy):
            logprint("截图成功,尝试了"+str(FailedCount)+"次", mainlog)
            FailedCount = 0
            BTfailed = False
            jumpWait = False
        else:
            FailedCount += 1
            continue
    elif not OCRR.get_img(lux, luy, rdx, rdy):  # 截图窗口
        BTfailed = True
        FailedCount += 1
        logprint("截图失败", mainlog)
        jumpWait = True
        continue"""
    OCRR.get_img_fwindow(hw)

    if np.array_equal(img_last, OCRR.img):  # 检测图像是否变化，无变化则跳过
        if jumpWait == False:
            logprint("图像未变化...", mainlog)
            jumpWait = True
        continue
    else:
        jumpWait = False
        img_last = OCRR.img
        OCRR.ocr()  # 对截图进行OCR
        if DEBUG_OCROT:
            logprint(str(OCRR), mainlog)  # 调试模式输出

    if OCRR.have("正在直播"):  # 查找到直播字符串
        res = OCRR.find("正在直播")  # 定位直播字符串
        time.sleep(1)
        if win32gui.IsIconic(hw):  # 钉钉设为前台
            win32gui.ShowWindow(hw, win32con.SW_SHOWNORMAL)

        if res_last[2] == res[2]:  # 如果直播内容没有改变
            continue
        else:  # 如果是新直播
            res_last = res
            click([lux+res[0], luy+res[1]])  # 点击字符串
            logprint("clicked " + str(res), mainlog)
    else:
        res_last = [0, 0, ""]
