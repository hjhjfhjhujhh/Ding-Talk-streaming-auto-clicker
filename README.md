# Ding-Talk-streaming-auto-clicker
自动点开钉钉直播，简单好用。基于python，opencv图像识别。

本代码开源遵循GPL3.0协议
# 注意事项：
- 本程序适配了多显示器的情况。
- 钉钉窗口无法最小化（不然得不到窗口截图），不过窗口无需置于最顶层，也就是说你可以把钉钉放在其他窗口后面。
- 这个脚本没有采用任何钉钉API，纯粹基于截图和鼠标点击。这意味着程序执行时可能会改变您的聚焦窗口，弹出来的直播窗口也可能将您的窗口覆盖。如果您正在玩游戏，这可能会对您的体验产生影响。
- 脚本不能处理多群直播同时开的情况，也不能处理签到，因为我暂时没有这个需求。有人需要的话可以给我提issue

本程序只有单个exe文件，打开之后全自动运行，使用方便快捷。

本程序在尽量实时的同时资源占用量极低，可有效保护您电脑的机体健康。

我还写过用PaddleOCR识别文字的版本。那个脚本我写了几个星期将近四百行，因为ppocr过于答辩的cpu占用、处理时间以及打包不能，我花了一天重写了这个cv2识别图片的版本。

# 用到的第三方模组
- numpy
- opencv-python
- pyqt5 (用来截图)
- qimage2ndarray

# 如何使用
- **如果您没有使用python以及GitHub的经验，请点击右方的Releases，下载下面的exe文件。**
- 如果您要直接使用源码，下载源码之后把resource里面的两张图片放在源码同目录，直接运行即可。


# 如何打包
- 首先安装pyinstaller：
- ```pip install pyinstaller```
- 把资源放在源码目录，在此目录打开命令行
- 把两张模板和图片图标打包进去：
- ```pyinstaller.exe .\DTalkClicker.py -F -i="logo1.ico" --add-data "stream_white.jpg;." --add-data "stream_black.jpg;."```
