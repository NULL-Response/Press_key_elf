import win32gui, win32api, win32con, win32ui
import time
from time import sleep
import random
from random import uniform
import logging
from logging import info, warning, error
from config import VK_CODE
import copy
import cv2 as cv
import traceback

# mode为test意为调试模式，real则为实际应用
# In 'test' mode, it will print some hidden info for Debug
mode = 'test'

# 模式图的存储路径
# The save path of pattern pics
pattern_path = 'yourpath'

# 设置logging的书写格式
# Setting logging configs
logging.basicConfig(
    level=logging.DEBUG,
    format=
    '%(lineno)d : %(asctime)s : %(levelname)s : %(funcName)s : %(message)s',
    filename='yourpath\\log{}.txt'.format(
        time.strftime('-%Y-%m-%d')),
    filemode='w')

# 用当前时间初始化随机数
# Set random seed
random.seed(time.time())


def Get_PosAndHwnd(hwnd=0, ClassName=None, TitleName=None):
    '''接受句柄、类名、标题名，返回第一个匹配的窗口位置和句柄。
        Accepts Handle, ClassName or TitleName of a window, returns the rect-pos of the first window that matches.
        
        Args:
            hwnd: 目标句柄 target handle.
            ClassName: 目标窗口类名 target ClassName.
            TitleName: 目标窗口标题 target TitleName.
        
        Returns: x1, y1, x2, y2, hwnd
            前四个参数描述窗口在显示器中的位置，hwnd为句柄。
            x1-y2 describes the pos of the window in screen, last arg is the handle.
    '''
    try:
        if hwnd == 0:
            hwnd = win32gui.FindWindow(ClassName, TitleName)
        x1, y1, x2, y2 = win32gui.GetWindowRect(hwnd)
        info('找到窗口，句柄：{}，标题名：{}，类名：{}，坐标：{}'.format(
            hwnd, win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd),
            (x1, y1, x2, y2)))
        return x1, y1, x2, y2, hwnd
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('Get_PosAndHwnd 出现错误')


def Activate_Hwnd(hwnd):
    '''激活指定句柄的窗口。每次转换时使用。
        Activates the window with the matching hwnd.

        Args:
            hwnd: 目标句柄 target handle.
        
        Returns:
            No return.
    '''
    try:
        win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        # WA_CLICKACTIVE为通过鼠标激活，WA_ACTIVE为鼠标以外的东西（如键盘）激活，WA_INACTIVE为取消激活
        # WA_CLICKACTIVE means activate by mouse, WA_ACTIVE means other equipments( like keyboard).WA_INACTIVE means inactivate it.
        info('激活句柄为 {} 标题为 {} 的窗口'.format(hwnd, win32gui.GetWindowText(hwnd)))
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('激活窗口 出现错误')


def LeftClick(hwnd, x1, y1, times='first', platform='PC'):
    '''左键单击（自带延时） Left Click the window (located by hwnd) once (with human-like delay). 

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            x1, y1: 需要点击的坐标 target pos. 
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
            platform: 目前无用。'PC'表示当前操作的是桌面版，不是安卓模拟器('AM')。
                Useless now. 'PC' means target window is PC, not Android manipulator ('AM').

        Returns:
            No return.
    '''
    try:
        if platform == 'PC':  # 对于桌面版 # for PC
            tmp_pos = win32api.MAKELONG(x1, y1)
        #elif flag == 'AM': # 对于模拟器 # for Android manipulator
        #click_pos = win32gui.ScreenToClient(hwnd, (x1, y1))
        # 注意坐标是整个窗口还是客户区的，用ScreenToClient转换 # switch between screen/client pos
        #tmp_pos = win32api.MAKELONG(click_pos[0], click_pos[1])
        if times == 'first':
            Activate_Hwnd(hwnd)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN,
                             win32con.MK_LBUTTON, tmp_pos)
        # 鼠标点击格式为SendMessage(hWnd,WM_LBUTTONDOWN,fwKeys,MAKELONG(xPos,yPos));
        # fwKeys可以取：MK_CONTROL、MK_LBUTTON、MK_MBUTTON、MK_RBUTTON、MK_SHIFT等，多个值用|隔开
        sleep(uniform(0.05, 0.08))
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0000, tmp_pos)
        # 抬起时有些电脑上fwKeys显示为0000，有些可能为MK_LBUTTON
        # Use spy++ for details
        sleep(uniform(0.07, 0.095))
        info('对句柄 {} 标题 {} 的坐标 {} 进行 左键单击 动作'.format(
            hwnd, win32gui.GetWindowText(hwnd), (x1, y1)))
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('左键单击 出现错误')


def LeftClick_sequence(hwnd, points, times='second', platform='PC'):
    '''左键依序单击（依次点击多个点） Click a sequence of points in order.

        Args: 
            hwnd: 目标窗口的句柄 target hwnd. 
            points: 需要点击的坐标 target points in the form like ((x1,y1),(x2,y2)...). 
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
            platform: 目前无用。'PC'表示当前操作的是桌面版，不是安卓模拟器('AM')。
                Useless now. 'PC' means target window is PC, not Android manipulator ('AM').
        
        Returns:
            No return.
    '''
    try:
        if times == 'first':
            Activate_Hwnd(hwnd)
        for point in points:
            LeftClick(hwnd, point[0], point[1], 'second', platform)
        info('对句柄 {} 标题 {} 的坐标 {} 进行 左键依序单击 动作'.format(
            hwnd, win32gui.GetWindowText(hwnd), points))
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('左键单击 出现错误')


def LeftDoubleClick(hwnd, x1, y1, times='first', platform='PC'):
    '''左键同一位置双击 Click the same pos twice.

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            x1, y1: 需要点击的坐标 target pos. 
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
            platform: 目前无用。'PC'表示当前操作的是桌面版，不是安卓模拟器('AM')。
                Useless now. 'PC' means target window is PC, not Android manipulator ('AM').

        Returns:
            No return.
    '''
    try:
        if times == 'first':
            Activate_Hwnd(hwnd)
        info('左键同一位置双击--第一次点击')
        LeftClick(hwnd, x1, y1, 'second', platform)
        info('左键同一位置双击--第二次点击')
        LeftClick(hwnd, x1, y1, 'second', platform)
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('左键同一位置双击 出现错误')


def LeftDifDoubleClick_Rect(hwnd,
                            x1,
                            y1,
                            times='first',
                            platform='PC',
                            delta=(0, 4)):
    '''左键不同位置双击，范围为双正方形 Left click 2 pos, in a double-square-like area.
        _________
        |  ___  |
        | |   | |
        | |___| |
        |_______|
        (x1, y1) is at the center, (x2, y2) between the squares (borders included).

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            x1, y1: 第一次点击的坐标 first pos to click. 
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
            platform: 目前无用。'PC'表示当前操作的是桌面版，不是安卓模拟器('AM')。
                Useless now. 'PC' means target window is PC, not Android manipulator ('AM').
            delta: 第二个点的横坐标在x1+delta[0] 到 x1+delta[1]间（包括端点），纵坐标同样。
                x2 varies between x1+delta[0] to x1+delta[1], y2 the same (both ends included). 

        Returns:
            No return.
    '''
    try:
        if times == 'first':
            Activate_Hwnd(hwnd)
        info('左键正方形双击--第一次点击')
        LeftClick(hwnd, x1, y1, 'second', platform)
        point_list = []
        for delta_x in range(delta[0], delta[1] + 1):
            for delta_y in range(delta[0], delta[1] + 1):
                point_list.append((x1 + delta_x, y1 + delta_y))
                point_list.append((x1 + delta_x, y1 - delta_y))
                point_list.append((x1 - delta_x, y1 + delta_y))
                point_list.append((x1 - delta_x, y1 - delta_y))
        point_list = list(set(point_list))
        point_list.remove((x1, y1))
        x1, y1 = random.sample(point_list, 1)[0]
        info('左键正方形双击--第二次点击')
        LeftClick(hwnd, x1, y1, 'second', platform)
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('左键正方形双击 出现错误')


def LeftDifDoubleClick_Cir(hwnd, x1, y1, times='first', platform='PC', r=4):
    '''左键不同位置双击，范围为圆形。 Left Click 2 pos in a circle.
        (x1, y1)为圆心，r为半径，(x2, y2)在圆内及圆上。
        (x1, y1) is the center, r the radius. Generates (x2, y2) in the circle (or on the border).

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            x1, y1: 第一次点击的坐标 first pos to click. 
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
            platform: 目前无用。'PC'表示当前操作的是桌面版，不是安卓模拟器('AM')。
                Useless now. 'PC' means target window is PC, not Android manipulator ('AM').
            r: 圆的半径。
                radius. 
        
        Returns:
            返回两次点击的坐标(x1,y1,x2,y2)。
            Returns 2 clicked pos in (x1,y1,x2,y2). 
    '''
    try:
        if times == 'first':
            Activate_Hwnd(hwnd)
        info('左键圆形双击--第一次点击')
        LeftClick(hwnd, x1, y1, 'second', platform)
        m, n = x1, y1
        point_list = []
        final_point_list = []
        for delta_x in range(r + 1):
            for delta_y in range(r + 1):
                point_list.append((delta_x, delta_y))
        point_list.remove((0, 0))
        for i in range(len(point_list)):
            x = point_list[i][0]
            y = point_list[i][1]
            if x * x + y * y <= r * r:
                final_point_list.append((x1 + x, y1 + y))
                final_point_list.append((x1 + x, y1 - y))
                final_point_list.append((x1 - x, y1 + y))
                final_point_list.append((x1 - x, y1 - y))
        final_point_list = list(set(final_point_list))
        x1, y1 = random.sample(final_point_list, 1)[0]
        info('左键圆形双击--第二次点击')
        LeftClick(hwnd, x1, y1, 'second', platform)
        return (m, n, x1, y1)
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('左键圆形双击 出现错误')


def Random_Areas_LeftClicks(hwnd,
                            click_areas,
                            times='first',
                            platform='PC',
                            r=4,
                            click_strategy=1):
    '''从几种点击策略中选一种执行，在一定范围对目标窗口进行点击。
        Select a click strategy/combination, click the window in some given areas.  
        ---------   单个点击区域如右。A single click_area looks like left. 
        |  ---- |   如果只点击一次或者点击两次中的第一个坐标，该坐标位于小矩形中（包括边界）。
        |  |  | |   If only click once, or the first click in 2 clicks, it locates in the minor rect (borders included).
        |  |  | |   小矩形的每条边距大矩形距离为r个像素点。
        |  |  | |   Each border of the minor rect is r-pixel far from the outside rect.
        |  ---- |
        ---------
        第二次点击的区域不会超出大矩形的范围（包括边界)）。对于一个点击区域[x1,y1,x2,y2]，其中的小矩形为[x1+r,y1+r,x2-r,y2-r]。
        The second is in the big rect (borders included). For a click_area like (x1,y1,x2,y2), the minor rect is [x1+r,y1+r,x2-r,y2-r]. 

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            click_areas: 可选的点击范围（矩形）。格式为[[x1,y1,x2,y2],[x3,y3,x4,y4]...]，只有一个区域则可以写成[x1,y1,x2,y2]。
                Possible rect-areas to click, given in the form of [[x1,y1,x2,y2],[x3,y3,x4,y4]...]. 
                If only one area, [x1,y1,x2,y2] is also acceptable. 
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
            platform: 目前无用。'PC'表示当前操作的是桌面版，不是安卓模拟器('AM')。
                Useless now. 'PC' means target window is PC, not Android manipulator ('AM').
            r: 圆的半径。
                radius. 
            click_strategy: 选择的点击策略。看看源码，很容易写自己的策略。主要包括点击方式和概率。
                A strategy for clicking, includes combination of clicking functions and its weights. 
                Look at the src code, easy to construct your own click_strategy. 
        
        Returns:
            返回两次点击的坐标(x1,y1,x2,y2)。如果只点击一次，返回(x1, y1, -1, -1)
            Returns 2 clicked pos in (x1,y1,x2,y2). If only clicked once, return (x1, y1, -1, -1). 
    '''
    try:
        # 缩小边界，并加上一个值，表示区域点的坐标个数（包含边界）
        total = 0
        areas = copy.deepcopy(
            click_areas
        )  # python中传列表参是浅拷贝，为了防止改变原click_areas需要写这一句。传字符串、数字是深拷贝
        if isinstance(areas[0], int):
            areas = [areas]
        for i in range(len(areas)):
            areas[i][0] = areas[i][0] + r
            areas[i][1] = areas[i][1] + r
            areas[i][2] = areas[i][2] - r
            areas[i][3] = areas[i][3] - r
            areas[i].append((areas[i][2] - areas[i][0] + 1) *
                            (areas[i][3] - areas[i][1] + 1))
            total = total + areas[i][4]
        rand = random.randint(0, total - 1)
        # 找出第一次点击位置落在哪个区域，并找出具体坐标
        # find out the first click is in which area, and the exact pos.
        x = 0
        y = 0
        for i in range(len(areas)):
            rand = rand - areas[i][4]
            if rand < 0:
                rand = rand + areas[i][4]
                w = areas[i][2] - areas[i][0] + 1
                x = rand % w + areas[i][0]
                y = rand // w + areas[i][1]  # 现在python中整型之间用/也会得到float
                break
        if click_strategy == 1:  # 1号策略，包括左键单击、左键同一位置双击、左键圆形双击三种
            #总结果数为173，左键单击:左键同一位置双击:左键圆形双击 的比例为 131:29:13
            rand = random.randint(1, 173)
            if times == 'first':
                Activate_Hwnd(hwnd)
            if rand <= 131:  # weighs 131
                info('随机区域左键 选择 左键单击，坐标{}'.format((x, y)))
                LeftClick(hwnd, x, y, 'second', platform)
                return (x, y, -1, -1)
            elif rand >= 161:  # weighs 13
                info('随机区域左键 选择 左键圆形双击，第一次点击坐标{}'.format((x, y)))
                return LeftDifDoubleClick_Cir(hwnd, x, y, 'second', platform,
                                              r)
            else:  # weighs 29
                info('随机区域左键 选择 左键同位置双击，坐标{}'.format((x, y)))
                LeftDoubleClick(hwnd, x, y, 'second', platform)
                return (x, y, x, y)
        elif click_strategy == 2:  # 2号策略，仅包括左键单击，用于只能单击的情况
            if times == 'first':
                Activate_Hwnd(hwnd)
            info('随机区域左键 选择 左键单击，坐标{}'.format((x, y)))
            LeftClick(hwnd, x, y, 'second', platform)
            return (x, y, -1, -1)
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('随机区域左键 出现错误')


def press_key(hwnd, key='esc', times='second'):
    '''单次按键（自带延时） Press a key once (with human-like delay).

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            key: 需要按的键，详见config.VK_CODE。
                The key to press, see also in config.VK_CODE.
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
        
        Returns:
            No return. 

    '''
    try:
        if times == 'first':
            Activate_Hwnd(hwnd)
        win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, VK_CODE[key],
                             0)  # 暂时将lParam写为0
        # 这里不知道是否要暂停
        # I don't konw whether to pause here.
        time.sleep(0.01)
        win32api.SendMessage(hwnd, win32con.WM_CHAR, VK_CODE[key], 0)
        sleep(uniform(0.05, 0.08))
        win32api.SendMessage(hwnd, win32con.WM_KEYUP, VK_CODE[key], 0)
        sleep(uniform(0.07, 0.095))
        info('对句柄 {} 标题 {} 的窗口 按 {} 键一次'.format(hwnd,
                                                win32gui.GetWindowText(hwnd),
                                                key))
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('单次按键 出现错误')


def press_keys(hwnd, keys, times='second'):
    '''依序按键 Press a sequence of keys in order.

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            keys: 需要按的键，详见config.VK_CODE。
                Keys to press, see also in config.VK_CODE.
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
        
        Returns:
            No return. 
    '''
    try:
        if times == 'first':
            Activate_Hwnd(hwnd)
        for key in keys:
            press_key(hwnd, key, 'second')
        info('对句柄 {} 标题 {} 的窗口 依序按 {} 键'.format(hwnd,
                                                win32gui.GetWindowText(hwnd),
                                                keys))
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('依序按键 出现错误')


def ClientRect_PrtSc(hwnd, area=None, filename=''):
    '''根据句柄、截图位置和图片路径，对窗口的客户区截图并存到指定位置。
        Press PrtSc for a window's client area, save the pic in the indicated path. 
        从https://blog.csdn.net/zhuisui_woxin/article/details/84345036和https://www.cnblogs.com/Evan-fanfan/p/11097850.html改编
        Made some changes and corrected some errors from 2 links above (took me some time to figure the errors).

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            area: 需要的区域。为None时表示整个客户区。
                Needed area. If None, means get the whole client area. 
            filename: 指定路径。为''时表示文件名根据hwnd生成。
                Indicate the path. If '', means 'path\{}.bmp'.format.  
        
        Returns:
            No return. 
    '''
    try:
        hwnd = hwnd
        if filename == '':
            filename = 'yourpath\\{}.bmp'.format(
                hwnd)
        hwndDC = win32gui.GetDC(
            hwnd)  # 获取窗口的设备上下文Device Context。GetWindowDC包括了非客户区，而GetDC仅为客户区
        # GetWindowDC also gets the non-client area, while GetDC only gets the client area.
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)  # 获取mfcDC
        saveDC = mfcDC.CreateCompatibleDC()  # 创建可兼容DC
        saveBitMap = win32ui.CreateBitmap()  # 创建bitmap以保存图片
        '''MonitorDev = win32api.EnumDisplayMonitors(
            None, None)  # 获取显示器信息，枚举显示器，笔记本据说可能有问题'''
        x1, y1, x2, y2 = win32gui.GetClientRect(
            hwnd)  #GetClientRect获取客户区窗口位置，GetWindowRect获取整个窗口的位置信息
        x, y, w, h = (0, 0, 0, 0)
        if area == None:
            x = 0
            y = 0
            w = x2 - x1
            h = y2 - y1
        else:
            x, y, m, n = area
            w = m - x
            h = n - y
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)  # 为bitmap开辟空间
        # 对saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)的理解：
        # 1.mfc相当于一个虚拟屏幕。这里的参数w和h决定了这个屏幕的大小。
        # 2.屏幕的初始状态是黑色，每个坐标都是#000000
        # 3.之前有mfcDC = win32ui.CreateDCFromHandle(hwndDC)，又有hwndDC = win32gui.GetDC(hwnd)
        #   mfcDC和hwnd窗口之间建立了某种关联，可以将hwnd窗口中的图像放到虚拟屏幕上
        saveDC.SelectObject(saveBitMap)  # 将截图保存到saveBitMap中
        saveDC.BitBlt((0, 0), (w, h), mfcDC, (x, y), win32con.SRCCOPY)
        # 对saveDC.BitBlt(坐标1, (w, h), mfcDC, 坐标2, win32con.SRCCOPY)的理解：
        # BitBlt的功能大概是把从hwnd窗口截到的图放到虚拟屏幕上，信息转入saveDC。
        # 1.坐标1是针对窗口截图的，指定截图放在黑色背景上的位置（指定左上角）
        # 2.w和h窗口截图的长宽，而坐标2指定了开始截图的位置
        #   这两个参数决定了从hwnd窗口的哪里截图、截多大的图
        # 3.mfcDC已经和hwnd窗口建立了关联，所以不需要指定虚拟屏幕从哪个窗口获得截图
        # 4.SRCCOPY意为将截图直接拷贝到虚拟屏幕中
        # 接下来的saveBitMap.SaveBitmapFile(saveDC, filename)则是对虚拟屏幕截图并保存到指定位置
        # Google for more info.
        saveBitMap.SaveBitmapFile(saveDC, filename)
        # 清除数据
        # GetDC一类的需要用ReleaseDC释放，CreateDC一类的用DeleteDC释放，DeleteObject则删除一个逻辑笔、画笔、字体、位图、区域或者调色板
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        info('对句柄 {} 标题 {} 的窗口截图并保存'.format(hwnd,
                                            win32gui.GetWindowText(hwnd)))
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('客户区截图 出现错误')


def G(pattern_filename):
    '''将pattern的文件名加上路径前缀与类型后缀并返回
        Add path prefix and '.bmp' for a pattern pic file and return it. 

        Args:
            pattern_filename: 模式图片名。Pattern pic file name. 
        
        Returns: 
            opencv2 可用的路径。The usable path for opencv2. 
    '''
    try:
        global pattern_path  # 在全局变量中 # a global variable.
        return pattern_path + pattern_filename + '.bmp'
    except:
        global mode
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('加模式图路径 出现错误')


def Find_Pic(hwnd,
             pattern,
             return_pos='l',
             find_pattern_in_area=None,
             filename=''):
    '''寻找图片，找到则返回小图在大图中的位置坐标及模板图大小，否则返回(-1,-1,tw,th)。
        Find the pattern pic in the target pic. Return the location and pattern's size if found, or (-1,-1,pattern-weight,pattern-height) if not.
        https://blog.csdn.net/wz2671/article/details/102751549和https://www.cnblogs.com/ssyfj/p/9271883.html很有帮助。
        Above 2 links help. 

        Args:
            hwnd: 目标窗口的句柄 target hwnd. 
            pattern: 模式图，小图。
                The pic to be found, pattern pic, smaller pic. 
            return_pos: 为'l'时返回找到的左上角坐标，'c'则返回中心（偏左上）。
                If 'l', return the left-top corner, 'c' the centor pos (if odd, lefter and topper). 
            find_pattern_in_area: 指定在大图的哪个区域查找。None表示在全部区域中找。
                In which area of the bigger pic to find. None means the whole area. 
            filename: 指定路径。为''时表示文件名根据hwnd生成。
                Indicate the bigger-pic's path. If '', means 'path\{}.bmp'.format.  
        
        Returns:
            如果找到，以(x, y, pat_w, pat_h)形式返回小图在大图中的位置；没找到则返回(-1, -1, pat_w, pat_h)。
            The location and pattern's size in the form of (x, y, pat_w, pat_h) if found, or (x, y, pat_w, pat_h) if not.
    '''
    try:
        hwnd = hwnd
        ClientRect_PrtSc(hwnd, find_pattern_in_area, filename)
        if filename == '':  # 没有提供文件名则需要改成默认文件名
            filename = 'yourpath\\{}.bmp'.format(
                hwnd)
        pat = cv.imread(pattern)  # pattern
        src = cv.imread(filename)  # source
        # imread的文件名必须带上格式后缀
        theight, twidth = pat.shape[:2]
        # 获取pattern的高和宽
        result = cv.matchTemplate(src, pat, cv.TM_SQDIFF_NORMED)
        # 执行模板匹配，采用的匹配方式cv2.TM_SQDIFF_NORMED
        # matchTemplate(大图像, 小图像(子图像、模式图像)m, 匹配方式, result 和 mask 可选)
        # 返回一个矩阵result，大小为(W-w+1)*(H-h+1)，比大图像少了右、下一圈，是子图像左上顶点在大图像中可能出现的位置
        # result矩阵中的每个位置，其值表示子图像左上顶点放在该位置时，子图像与大图像上相应区域的匹配度
        # result中的点位置和大图像保持一致。
        # 关于匹配方式：（）
        # 1.TM_SQDIFF是平方差匹配；TM_SQDIFF_NORMED是标准平方差匹配。利用平方差来进行匹配,最好匹配为0；匹配越差,匹配值越大。
        # 2.TM_CCORR是相关性匹配；TM_CCORR_NORMED是标准相关性匹配。采用模板和图像间的乘法操作,数越大表示匹配程度较高, 0表示最坏的匹配效果。
        # 3.TM_CCOEFF是相关性系数匹配；TM_CCOEFF_NORMED是标准相关性系数匹配。将模版对其均值的相对值与图像对其均值的相关值进行匹配,1表示完美匹配,-1表示糟糕的匹配,0表示没有任何相关性(随机序列)。
        # 总结：随着从简单的测量(平方差)到更复杂的测量(相关系数),我们可获得越来越准确的匹配(同时也意味着越来越大的计算代价)。
        min_val, max_val, min_loc, max_loc = cv.minMaxLoc(result)
        # minMaxLoc(矩阵Mat)，在一个矩阵中找出最大值、最小值、对应序号
        # 对于TM_SQDIFF_NORMED模式下最匹配的位置，如果不匹配值大于1%，就认为没有找到。
        if min_val > 0.01:
            return (-1, -1, twidth, theight)
        global mode
        if mode == 'test':
            # 以下代码绘制矩形边框，并将匹配区域标注出来，用于调试
            # codes after are used for debug, draw the big/small pic and the location.
            cv.rectangle(src, min_loc,
                         (min_loc[0] + twidth, min_loc[1] + theight),
                         (0, 0, 225), 1)  # 在src图像矩阵中加上矩形，范围是模板图的范围
            # rectangle(目标图像, 矩形的一个顶点, 矩形对角线上的另一个顶点, 矩形线条颜色, 线条粗细)返回空值
            strmin_val = str(min_val)
            print('\n匹配度' + strmin_val)
            cv.imshow("MatchResult----MatchingValue=" + strmin_val,
                      src)  # 将画上了矩形的图像在窗口中显示出来
            cv.imshow('pattern pic', pat)
            # imshow(窗口名称, 图像矩阵)，在窗口中显示图像，窗口自动调整到图像大小
            # cv.imwrite('1.png', template, [int(cv.IMWRITE_PNG_COMPRESSION), 9])
            # imwrite(文件名, 图像矩阵, 格式编码) 将图像矩阵以特定编码保存，暂时用不到
            # IMWRITE_PNG_COMPRESSION为png格式，0-9为压缩级别，9为不压缩，消耗时间最少
            cv.waitKey(0)  # 等待键盘输入任意值，参数表示等待时间（ms），0表示一直等
            cv.destroyAllWindows()  # 关闭所有窗口
        if return_pos == 'c':
            x = min_loc[0] + twidth // 2
            y = min_loc[1] + theight // 2
        elif return_pos == 'l':
            x = min_loc[0]
            y = min_loc[1]
        info('在句柄 {} 标题 {} 的窗口找到模式{}'.format(hwnd,
                                             win32gui.GetWindowText(hwnd),
                                             pattern))
        return (x, y, twidth, theight)
    except:
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('找图 出现错误')


def print_hitted_points(hitted_points, hwnd, deltas):
    '''测试用函数，根据hwnd找到截图文件（不要重新截图）并标注出点击过的点。
        Used for debug, find the big pic by hwnd (No PrtSc again) and mark the clicked pos.

        Args:
            hitted_points: 点击过的点。clicked points. 
            hwnd: 用于找到大图。Used to find the big pic.
            deltas: 如果大图是客户区的一部分截图，该变量可描述大图左上角顶点在整个客户区中的位置。
                If the big pic is only part of the window, deltas would help describe big-pic's location (deltas is the left-top corner pos). 
        
        Returns:
            No return. 
    '''
    try:
        x1, y1, x2, y2 = hitted_points
        x1, x2 = x1 - deltas[0], x2 - deltas[0]
        y1, y2 = y1 - deltas[1], y2 - deltas[1]
        print('deltas = {}, '.format(deltas))
        print('hitted_points = ({},{})  ({},{})'.format(x1, y1, x2, y2))
        # 这样就得到了应该标注的坐标
        filename = 'yourpath\\{}.bmp'.format(
            hwnd)
        src = cv.imread(filename)
        if x2 < 0 and y2 < 0:
            cv.circle(src, (x1, y1), 1, (0, 0, 225), 4)
            # cv.circle(img, point, point_size, point_color, thickness)本来用来画圆的
            # 参数分别是图像矩阵、圆心、半径、线的颜色、稠度
            # 稠度可以取0、4、8，画出来点的大小不同
            cv.imshow('single one----hitted points', src)
        elif x1 == x2 and y1 == y2:
            cv.circle(src, (x1, y1), 1, (0, 0, 225), 4)
            cv.imshow('double same----hitted points', src)
        else:
            cv.circle(src, (x1, y1), 1, (0, 0, 225), 4)
            cv.circle(src, (x2, y2), 1, (0, 0, 225), 4)
            cv.imshow('double diff----hitted points', src)
        cv.waitKey(0)
        cv.destroyAllWindows()
    except:
        traceback.print_exc()


def FindPic_RandomLeftClick(hwnd,
                            pattern,
                            max_try=20,
                            time_interval=0.5,
                            click_strategy=1,
                            times='first',
                            click_areas=None,
                            find_pattern_in_area=None,
                            r=4,
                            platform='PC',
                            click_pos='l',
                            filename=''):
    '''从指定的句柄窗口寻找图片，若最大次数内找到则执行随机区域左键点击，否则返回False。返回点击的坐标(x1,y1,x2,y2)，若只点击了一次则返回(x1,y1,-1,-1)。
        这是一个整合程度高函数，可能要看一下其中用到的其他函数。
        Find a pattern pic in the hwnd window, if not found after max_try, return False; if found, operate a click_strategy, and return clicked pos. 
        If clicked only once, return (x1,y1,-1,-1). 
        This func is highly intergrated. Turn to some called sub-funcs if confused. 

        Args in detail:
            hwnd: 要截图并点击的窗口句柄。The window to PrtSc and click. 
            pattern: 要寻找的模式小图。The pattern-pic's path to be found. 
            max_try: 指定最多寻找的次数。If not found, after max_try, it will let go. 
            time_interval指定每次寻找的间隔时间（实际查找间隔比该时间多一点）。
                Interval between each try (a little longer than given, in fact).
            click_strategy表示点击策略，详见Random_Areas_LeftClicks函数。
                Indicate the click strategy, see also in Random_Areas_LeftClicks(). 
            times: 'first'表示从其他窗口转过来点击，用于调用Activate_Hwnd(hwnd)。
                If 'first', means target window is inactive/unfocused, thus call Activate_Hwnd(hwnd). 
            click_areas: 可选的点击范围（矩形）。格式为[[x1,y1,x2,y2],[x3,y3,x4,y4]...]，只有一个区域则可以写成[x1,y1,x2,y2]。为None时表示就点模式小图所在的区域。
                Possible rect-areas to click, given in the form of [[x1,y1,x2,y2],[x3,y3,x4,y4]...]. 
                If only one area, [x1,y1,x2,y2] is also acceptable. 
                If None, means to click the area where pattern pic is found. 
            find_pattern_in_area: 指定在窗口客户区的哪个区域寻找模式图，为None时表示在整个客户区寻找。
                Which rect-area of the client area of the window to find the pattern pic. If None, means find in the whole area of the client area.
                find_pattern_in_area从技术上是只截取客户区该区域的图实现的。
                If not None, technically, it works by only PrtSc the partly area of the window's client area. 
            r: 指示对指定的点击区域，应该内收多少个像素点。对于高或宽<4的模式图，可能需要自己把r设置成更小的值。
                Shrink r pixels for each click_area. For pat-pic smaller than 4 pixels, adjust r to a smaller value. 
            platform: 目前无用。'PC'表示当前操作的是桌面版，不是安卓模拟器('AM')。
                Useless now. 'PC' means target window is PC, not Android manipulator ('AM').
            click_pos: 可取值'l'或'c'，l表示left_top，后者表示center。详见Find_Pic()。
                Possible values are 'l' and 'c'. See also in Find_Pic(). 
                用于指示在窗口中找到模式图后，是返回其左上坐标还是中心点坐标（模式图长宽为奇数时返回值偏左上）。
                'l' means left-top corner pos, 'c' means center pos, used by Find_Pic(). 
                现在暂时只用'l'模式，'c'模式可能用于以后功能拓展。
                For now 'l' is enough, 'c' may help expand furthur functions. 
            filename: 手动指定大图文件，从该文件中查找模式图。一般用不到。
                Indicate the big-pic's path manually, not used in ordinary applications.  
        
        Returns:
            若最大次数后仍未找到模式图则返回False。
            若找到模式图，返回点击的坐标(x1,y1,x2,y2)，若只点击了一次则返回(x1,y1,-1,-1)。
            If pattern pic not found after max_try, return False;
            if found, return clicked pos in (x1,y1,x2,y2). 
            Return (x1,y1,-1,-1) if only clicked once. 
    '''
    try:
        global mode
        x, y = -1, -1
        deltas = (
            -1, -1
        )  # 这是截图区域与整个客户区的左上顶点的相对位置。用于测试 # Used for debug, in print_hitted_points().
        for i in range(max_try):
            x, y, twidth, theight = Find_Pic(hwnd, pattern, 'l',
                                             find_pattern_in_area, filename)
            if x == -1 and y == -1:
                if mode == 'test':
                    print('\r第{}次没找到'.format(i + 1), end='')
                sleep(time_interval)
                continue
            else:
                if find_pattern_in_area == None:  # 若使用默认的全客户区查找，返回的坐标是相对于整个客户区的
                    delta_x, delta_y = 0, 0
                else:  # 指定了查找区域，则返回的坐标是相对于查找区域的，需要转换一下
                    delta_x, delta_y = find_pattern_in_area[
                        0], find_pattern_in_area[1]
                if click_areas == None:  # 不给出点击区域，则默认点击区域是窗口找到模式图片的区域
                    click_areas = [[
                        x + delta_x, y + delta_y, x + twidth + delta_x,
                        y + theight + delta_y
                    ]]
                if mode == 'test':
                    deltas = (delta_x, delta_y)
                break
        if x == -1 and y == -1:
            info('未能找到并随机点击：对句柄 {} 标题 {} 的窗口未找到模式{}，等待了{}次*{}秒'.format(
                hwnd, win32gui.GetWindowText(hwnd), pattern, max_try,
                time_interval))
            return False
        else:
            if mode == 'test':  # 在测试模式下，画出点击了窗口上的哪些点
                hitted_points = Random_Areas_LeftClicks(
                    hwnd, click_areas, times, platform, r, click_strategy)
                print_hitted_points(hitted_points, hwnd, deltas)
                return hitted_points
            elif mode == 'real':
                return Random_Areas_LeftClicks(hwnd, click_areas, times,
                                               platform, r, click_strategy)
    except:
        if mode == 'test':
            traceback.print_exc()
        elif mode == 'real':
            error('找图并随机左键点击 出现错误')


if __name__ == '__main__':
    try:
        click_areas = [292, 175, 314, 192]
        find_pattern_in_area = [44, 82, 245, 313]  # [44, 82, 100, 150]
        print(FindPic_RandomLeftClick(3212840, G('pattern1'),
                                      click_strategy=2))
    except:
        traceback.print_exc()
