import random
import win32api,win32gui
import sys
import time
import win32con
import logging
from PIL import ImageGrab, Image
from ssim.ssimlib import SSIM, SSIMImage
from ssim.utils import get_gaussian_kernel
import pdb
import time
import pyautogui

logging.basicConfig(level=logging.INFO,
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
    win32gui.SetForegroundWindow(hwnd)
    logging.debug("窗口坐标：")
    logging.debug(str(left)+' '+str(right)+' '+str(top)+' '+str(bottom))
else:
    logging.debug("未找到窗口")
    sys.exit(0)

#锁定雷区坐标
left += 38
top += 80
# right: 238 - (38 + 9 * 18)
right -= 38
# bottom: 283 - (80 + 9 * 18)
bottom -= 41

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

#数字1-8
#0 未被打开 img_0
#-1 被打开 空白, img_b
#-2 问号 img_q
#-4 红旗 img_f
#-8 炸弹 img_b0, img_b1, img_b2
img_0=Image.open("bmp/0.bmp").crop((3,3,14,14))
#img_0a=Image.open("bmp/0a.bmp").crop((3,3,14,14))
img_b=Image.open("bmp/b.bmp").crop((3,3,14,14))
img_f=Image.open("bmp/f.bmp").crop((3,3,14,14))
img_1=Image.open("bmp/1.bmp").crop((3,3,14,14))
img_2=Image.open("bmp/2.bmp").crop((3,3,14,14))
img_3=Image.open("bmp/3.bmp").crop((3,3,14,14))
img_4=Image.open("bmp/4.bmp").crop((3,3,14,14))
img_5=Image.open("bmp/5.bmp").crop((3,3,14,14))
img_6=Image.open("bmp/6.bmp").crop((3,3,14,14))
img_7=Image.open("bmp/7.bmp").crop((3,3,14,14))
img_8=Image.open("bmp/8.bmp").crop((3,3,14,14))
img_b0=Image.open("bmp/bomb0.bmp").crop((3,3,14,14))
img_b1=Image.open("bmp/bomb1.bmp").crop((3,3,14,14))
img_b2=Image.open("bmp/bomb2.bmp").crop((3,3,14,14))
img_q=Image.open("bmp/q.bmp").crop((3,3,14,14))

compare_threshold=0.7

tup_images = (
    img_0,
#    img_0a,
    img_b,
    img_f,
    img_1,
    img_2,
    img_3,
    img_4,
    img_5,
    img_6,
    img_7,
    img_8,
    img_b0,
    img_b1,
    img_b2,
    img_q
)

tup_code = (0,-1,-4, 1, 2, 3, 4, 5, 6, 7, 8, -8, -8, -8, -2)

def compare_image(this_image):
    "识别图像"
    ret = 9
    #ret = 0
    gaussian_kernel_sigma = 1.5
    gaussian_kernel_width = 11
    gaussian_kernel_1d = get_gaussian_kernel(
        gaussian_kernel_width, gaussian_kernel_sigma) 
    size = None   
    ssim = SSIM(this_image, gaussian_kernel_1d, size=size)
    #ssim_value = ssim.ssim_value(img_0)
    count = 0
    for comparison_image in tup_images:
        #comparison_image = comparison_image.crop((3,3,14,14))
        ssim_value = ssim.ssim_value(comparison_image)
        #print(count, ssim_value)
        
        if ssim_value > compare_threshold:
            #print(count, ssim_value, tup_code[count])
            ret = tup_code[count]
            break

        count = count + 1
    return ret

def printmap():
    for y in range(blocks_y):
        for x in range(blocks_x):
            print(map[y][x], end=" ")
        print()


map = [[0 for i in range(blocks_x)] for i in range(blocks_y)]

#用于统计是否需要随机点击
luckly = 0
#是否游戏结束
gameover = 0

#扫描雷区图像
def showmap():
    print("扫描雷区图像")
    img = ImageGrab.grab().crop(rect)

    for y in range(blocks_y):
        for x in range(blocks_x):
            this_image = img.crop((x * block_width, y * block_height, (x + 1) * block_width, (y + 1) * block_height))
            #this_image.save("all/"+str(y)+"_"+str(x)+".bmp", "BMP")
            this_image = this_image.crop((3,3,14,14))
            map[y][x] = compare_image(this_image)
            if map[y][x] == -8 or map[y][x] == 9:
                global gameover
                gameover = 1
                printmap()
                exit(0)
                break

#插旗
def banner():
    print("插旗")
    showmap()
    for y in range(blocks_y):
        for x in range(blocks_x):
            if 1 <= map[y][x] and map[y][x] <= 8:
                boom_number = map[y][x]
                print("坐标为", y, x, "数字为", boom_number)
                #    pdb.set_trace()
                block_white = 0
                block_qi = 0
                #pdb.set_trace()
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
                #print("空白数为", block_white, "旗子数", block_qi, "炸弹数", boom_number)
                #pdb.set_trace()
                if boom_number == block_white + block_qi and block_white !=0:
                    print("有确定的,坐标为", y, x)
                    printmap()
                    for yy in range(y - 1, y + 2):
                        for xx in range(x - 1, x + 2):
                            if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                                if not (yy == y and xx == x):
                                    if map[yy][xx] == 0:
                                        print("插旗", yy, xx)
                                        #pdb.set_trace()
                                        logging.debug("插旗" + str(yy) + "," + str(xx))
                                        #win32gui.SetForegroundWindow(hwnd)
                                        #win32api.SetCursorPos([left+xx*block_width, top+yy*block_height])
                                        # print("移动到")
                                        # print(left+xx*block_width)
                                        # print(top+yy*block_height)
                                        #win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                                        #time.sleep(0.05)
                                        #win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                                        #win32api.SetCursorPos([left, top])
                                        #time.sleep(0.09)
                                        pyautogui.click(button='right', x=left+xx*block_width+1, y=top+yy*block_height+1)
                                        pyautogui.moveTo(left, top)
                                        map[yy][xx] = -4
                                        #showmap()

#点击白块
def dig():
    print("点击白块")
    showmap()
    iscluck = 0
    for y in range(blocks_y):
        for x in range(blocks_x):
            if 1 <= map[y][x] and map[y][x] <= 8:
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
                                        print("点开", yy, xx)
                                        logging.debug("点开")
                                        logging.debug(yy)
                                        logging.debug(xx)
                                        #win32api.SetCursorPos([left + xx * block_width, top + yy * block_height])
                                        # print("移动到")
                                        # print(left + xx * block_width)
                                        # print(top + yy * block_height)
                                        #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                                        #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                                        pyautogui.click(left + xx * block_width + 1, top + yy * block_height + 1)
                                        pyautogui.moveTo(left, top)
                                        iscluck = 1
    if iscluck == 0:
        luck()

#随机点击
def luck():
    print("随机点击")
    fl = 1
    while(fl):
        random_x = random.randint(0, blocks_x - 1)
        random_y = random.randint(0, blocks_y - 1)
        if(map[random_y][random_x] == 0):
            print("乱点一个", random_y, random_x)
            logging.debug("乱点一个:" + str(random_y) + "," + str(random_x))
            #win32api.SetCursorPos([left + random_x * block_width, top + random_y * block_height])
            #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            pyautogui.click(left + random_x * block_width + 1, top + random_y * block_height + 1)
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
    #win32api.SetCursorPos([left, top])
    #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    pyautogui.click(left + 1, top + 1)
    showmap()
    global gameover
    while(1):
        if(gameover == 0):
            banner()
            #banner()
            dig()
        else:
            gameover = 0
            #win32api.keybd_event(113, 0, 0, 0)

            # Alt + N
            #win32api.keybd_event(0x12, 0, 0, 0)
            #win32api.keybd_event(78, 0, 0, 0)
            #win32api.keybd_event(78, 0, win32con.KEYEVENTF_KEYUP, 0)
            #win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)

            #win32api.SetCursorPos([left, top])
            #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            #win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            showmap()

gogo()
