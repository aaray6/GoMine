import win32gui
import win32con
import win32api
import sys
import getopt

from PIL import ImageGrab, Image

from ssim.ssimlib import SSIM, SSIMImage
from ssim.utils import get_gaussian_kernel

#扫雷游戏窗口
#class_name = "TMain"
class_name = "Minesweeper"
#title_name = "Minesweeper Arbiter "
title_name = "扫雷"
hwnd = win32gui.FindWindow(class_name, title_name)

#窗口坐标
left = 0
top = 0
right = 0
bottom = 0

if hwnd:
    print("找到窗口")
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    #win32gui.SetForegroundWindow(hwnd)
    print("窗口坐标：")
    print(str(left)+' '+str(right)+' '+str(top)+' '+str(bottom))
    #print("class_name:" + win32gui.GetClassName(hwnd))
else:
    print("未找到窗口")

### args ###
outputfile = "minesweeper00.bmp"

if(len(sys.argv) > 1):
    print(str(sys.argv))
    try:
        opts, args = getopt.getopt(sys.argv[1:],"ho:",["ofile="])
        print(str(opts))
    except getopt.GetoptError:
        print ("Usage:", sys.argv[0], '-o <outputfile>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ("Usage:", sys.argv[0], '-o <outputfile>')
            sys.exit()
        elif opt in ("-o", "--ofile"):
            outputfile = arg
    print ('Output file is', outputfile)
else:
    print("no args. the default output file is", outputfile)

#########

### 
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

img.save(outputfile, "BMP")

#每个方块18*18
block_width, block_height = 18, 18
#横向有blocks_x个方块
blocks_x = int((right - left) / block_width)
print("横向:" + str(blocks_x))
#纵向有blocks_y个方块
blocks_y = int((bottom - top) / block_height)
print("纵向:" + str(blocks_y))

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
#0 未被打开 img_0
#-1 被打开 空白, img_b
#-4 红旗 img_f
#-8 炸弹 img_b0, img_b1, img_b2
img_0=Image.open("bmp/0.bmp").crop((3,3,14,14))
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

compare_threshold=0.8

tup_images = (
    img_0,
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
    img_b2
)

tup_code = (0,-1,-4, 1, 2, 3, 4, 5, 6, 7, 8, -8, -8, -8)

def compare_image(this_image):
    "识别图像"
    ret = 9
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

map = [[0 for i in range(blocks_x)] for i in range(blocks_y)]


#扫描雷区图像
def showmap():
    img = ImageGrab.grab().crop(rect)
    #print(img.getcolors())
    
    #thresh = 200
    #fn = lambda x : 255 if x > thresh else 0
    #r = img.convert('L').point(fn, mode='1')
    #r.show()

    #img = img.convert("1")

    #height,width = img.size
    #image_data = img.load()
    #for loop1 in range(height):
    #    for loop2 in range(width):
    #        r,g,b = image_data[loop1,loop2]
    #        image_data[loop1,loop2] = r,0,0
    #img.show()

    #print(img.getcolors())
    for y in range(blocks_y):
        for x in range(blocks_x):
            this_image = img.crop((x * block_width, y * block_height, (x + 1) * block_width, (y + 1) * block_height))
            this_image = this_image.crop((3,3,14,14))
            #this_image.show()
            #logging.debug('x:' + str(x) + ' y:' +str(y) + ' colors:' + str(this_image.getcolors()))
            #print(this_image.getcolors())
            #grey_img = img.convert("L")
            #this_image = grey_img.crop((x * block_width, y * block_height, (x + 1) * block_width, (y + 1) * block_height))
            #logging.debug('x:' + str(x) + ' y:' +str(y) + ' colors:' + str(this_image.getcolors()))

            #grey_image.show()
            #sys.exit(0)
            #if(y==2 and x== 2):
            #    this_image.save("7.bmp", "BMP")
            #if(y==6 and x== 26):
            #    this_image.save("4.bmp", "BMP")
            #if(y==8 and x== 15):
            #    this_image.save("f.bmp", "BMP")
            #if(y==8 and x== 13):
            #    this_image.save("q.bmp", "BMP")
            #if(y==4 and x== 4):
            #    this_image.save("0_04_04.bmp", "BMP")
            ###this_image.save("all/"+str(y)+"_"+str(x)+".bmp", "BMP")

            """
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
                #print("无法识别图像")
                #print("坐标")
                #print((y,x))
                map[y][x] = -8
                #print("颜色")
                #print(this_image.getcolors())
                #sys.exit(0)
            """
            map[y][x] = compare_image(this_image)
            #print(map[y][x], end=" ")
        #print()
    #print(map)
    for y in range(blocks_y):
        for x in range(blocks_x):
            print(map[y][x], end=" ")
        print()

showmap()