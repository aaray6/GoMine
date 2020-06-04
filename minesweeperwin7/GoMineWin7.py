import random
import win32api,win32gui
import sys
import time
import win32con
from PIL import ImageGrab
import logging

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                    datefmt='%d %b %H:%M:%S')
#logging.basicConfig(level=logging.DEBUG,
#                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
#                    datefmt='%a, %d %b %Y %H:%M:%S',
#                    filename='/tmp/test.log',
#                    filemode='w')


logging.debug('开始')


#扫雷游戏窗口
class_name = "Minesweeper"
title_name = "扫雷"
hwnd = win32gui.FindWindow(class_name, title_name)

#窗口坐标
left = 0
top = 0
right = 0
bottom = 0

#初级窗口大小 0 w:238 0 h:283
#每个格子大小18x18,低级9x9格子

if hwnd:
    logging.debug("找到窗口")
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    #win32gui.SetForegroundWindow(hwnd)
    logging.debug("窗口坐标：")
    logging.debug(str(left)+' '+str(right)+' '+str(top)+' '+str(bottom))
else:
    logging.debug("未找到窗口")

#锁定雷区坐标
left += 39
top += 81
# right: 238 - (39 + 9 * 18)
right -= 37
# bottom: 283 - (81 + 9 * 18)
bottom -= 40

#抓取雷区图像
rect = (left, top, right, bottom)
img = ImageGrab.grab().crop(rect)

#每个方块18*18
block_width, block_height = 18, 18
#横向有blocks_x个方块
blocks_x = int((right - left) / block_width)
logging.debug("横向:" + str(blocks_x))
#纵向有blocks_y个方块
blocks_y = int((bottom - top) / block_height)
logging.debug("纵向:" + str(blocks_y))

rgba_ed = [(225, (192, 192, 192)), (31, (128, 128, 128))]
rgba_hongqi = [(54, (255, 255, 255)), (17, (255, 0, 0)), (109, (192, 192, 192)), (54, (128, 128, 128)), (22, (0, 0, 0))]
rgba_0 = [(54, (255, 255, 255)), (148, (192, 192, 192)), (54, (128, 128, 128))]
rgba_1 = [(185, (192, 192, 192)), (31, (128, 128, 128)), (40, (0, 0, 255))]
rgba_2 = [(160, (192, 192, 192)), (31, (128, 128, 128)), (65, (0, 128, 0))]
rgba_3 = [(62, (255, 0, 0)), (163, (192, 192, 192)), (31, (128, 128, 128))]
rgba_4 = [(169, (192, 192, 192)), (31, (128, 128, 128)), (56, (0, 0, 128))]
rgba_5 = [(70, (128, 0, 0)), (155, (192, 192, 192)), (31, (128, 128, 128))]
rgba_6 = [(153, (192, 192, 192)), (31, (128, 128, 128)), (72, (0, 128, 128))]
rgba_8 = [(149, (192, 192, 192)), (107, (128, 128, 128))]

rgba_boom = [(4, (255, 255, 255)), (144, (192, 192, 192)), (31, (128, 128, 128)), (77, (0, 0, 0))]
rgba_boom_red = [(4, (255, 255, 255)), (144, (255, 0, 0)), (31, (128, 128, 128)), (77, (0, 0, 0))]
#数字1-8
#0 未被打开
#-1 被打开 空白
#-4 红旗

map = [[0 for i in range(blocks_x)] for i in range(blocks_y)]

#用于统计是否需要随机点击
luckly = 0
#是否游戏结束
gameover = 0

#扫描雷区图像
def showmap():
    img = ImageGrab.grab().crop(rect)
    for y in range(blocks_y):
        for x in range(blocks_x):
            this_image = img.crop((x * block_width, y * block_height, (x + 1) * block_width, (y + 1) * block_height))
            if this_image.getcolors() == rgba_0:
                map[y][x] = 0
            elif this_image.getcolors() == rgba_1:
                map[y][x] = 1
            elif this_image.getcolors() == rgba_2:
                map[y][x] = 2
            elif this_image.getcolors() == rgba_3:
                map[y][x] = 3
            elif this_image.getcolors() == rgba_4:
                map[y][x] = 4
            elif this_image.getcolors() == rgba_5:
                map[y][x] = 5
            elif this_image.getcolors() == rgba_6:
                map[y][x] = 6
            elif this_image.getcolors() == rgba_8:
                map[y][x] = 8
            elif this_image.getcolors() == rgba_ed:
                map[y][x] = -1
            elif this_image.getcolors() == rgba_hongqi:
                map[y][x] = -4
            elif this_image.getcolors() == rgba_boom or this_image.getcolors() == rgba_boom_red:
                global gameover
                gameover = 1
                break
                #sys.exit(0)
            else:
                logging.debug("无法识别图像")
                logging.debug("坐标")
                logging.debug((y,x))
                logging.debug("颜色")
                logging.debug(this_image.getcolors())
                sys.exit(0)
    #print(map)

#插旗
def banner():
    showmap()
    for y in range(blocks_y):
        for x in range(blocks_x):
            if 1 <= map[y][x] and map[y][x] <= 5:
                boom_number = map[y][x]
                # print("数字为")
                # print(boom_number)
                # print("坐标为")
                # print(y)
                # print(x)
                block_white = 0
                block_qi = 0
                for yy in range(y-1,y+2):
                    for xx in range(x-1,x+2):
                        if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                            if not (yy == y and xx == x):
                                # print("map的值")
                                # print(map[yy][xx])
                                if map[yy][xx] == 0:
                                    block_white += 1
                                elif map[yy][xx] == -4:
                                    block_qi += 1
                # print("空白数为")
                # print(block_white)
                if boom_number == block_white + block_qi:
                    # print("有确定的,坐标为")
                    # print(y)
                    # print(x)
                    for yy in range(y - 1, y + 2):
                        for xx in range(x - 1, x + 2):
                            if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                                if not (yy == y and xx == x):
                                    if map[yy][xx] == 0:
                                        logging.debug("插旗")
                                        logging.debug(yy)
                                        logging.debug(xx)
                                        win32api.SetCursorPos([left+xx*block_width, top+yy*block_height])
                                        # print("移动到")
                                        # print(left+xx*block_width)
                                        # print(top+yy*block_height)
                                        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                                        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                                        showmap()

#点击白块
def dig():
    showmap()
    iscluck = 0
    for y in range(blocks_y):
        for x in range(blocks_x):
            if 1 <= map[y][x] and map[y][x] <= 5:
                boom_number = map[y][x]
                # print("数字为")
                # print(boom_number)
                # print("坐标为")
                # print(y)
                # print(x)
                block_white = 0
                block_qi = 0
                for yy in range(y - 1, y + 2):
                    for xx in range(x - 1, x + 2):
                        if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                            if not (yy == y and xx == x):
                                if map[yy][xx] == 0:
                                    block_white += 1
                                elif map[yy][xx] == -4:
                                    block_qi += 1
                # print("空白数为")
                # print(block_white)
                if boom_number == block_qi and block_white > 0:
                    # print("有确定的,坐标为")
                    # print(y)
                    # print(x)
                    for yy in range(y - 1, y + 2):
                        for xx in range(x - 1, x + 2):
                            if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                                if not(yy == y and xx == x):
                                    if map[yy][xx] == 0:
                                        logging.debug("点开")
                                        logging.debug(yy)
                                        logging.debug(xx)
                                        win32api.SetCursorPos([left + xx * block_width, top + yy * block_height])
                                        # print("移动到")
                                        # print(left + xx * block_width)
                                        # print(top + yy * block_height)
                                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                                        iscluck = 1
    if iscluck == 0:
        luck()

#随机点击
def luck():
    fl = 1
    while(fl):
        random_x = random.randint(0, blocks_x - 1)
        random_y = random.randint(0, blocks_y - 1)
        if(map[random_y][random_x] == 0):
            logging.debug("乱点一个")
            logging.debug(random_y)
            logging.debug(random_x)
            win32api.SetCursorPos([left + random_x * block_width, top + random_y * block_height])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            fl = 0

# showmap()
# win32api.SetCursorPos([left,top])
# win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
# win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
#time.sleep(1)
# showmap()
#banner()
#dig()

def gogo():
    win32api.SetCursorPos([left, top])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    showmap()
    global gameover
    while(1):
        if(gameover == 0):
            banner()
            banner()
            dig()
        else:
            gameover = 0
            win32api.keybd_event(113, 0, 0, 0)
            win32api.SetCursorPos([left, top])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            showmap()

gogo()
