
from selenium.webdriver import Edge
import time
import cv2
from PIL import ImageFont, ImageDraw, Image
import numpy as np
def search_shot(driver,search_comment,save_path):
    """
    输入搜索内容后进行搜索并截图保存在相应目录
    @param driver: selenium控制变量
    @param search_comment: 将要搜索的内容
    @param save_path: 图片截图保存路径
    @return: none
    """
    driver.find_element("xpath",'/html/body/div[2]/div[2]/input').send_keys(search_comment);        # 定位输入文本框并将搜索内容输入
    driver.find_element("xpath",'/html/body/div[2]/div[2]/div').click();                            # 定位按钮并点击搜索按键
    driver.get_screenshot_as_file(save_path);                                                       # 保存图片

def Add_Str(open_path,save_path,font,text):
    """
    在图片中加入规定文字后保存图片
    @param open_path: 打开图片的路径
    @param save_path: 保存图片的路径
    @param font: 字体
    @param text: 插入的文本
    @return: None
    """
    img=cv2.imread(open_path)

    img_pil = Image.fromarray(img)                                                                  #cv2格式图片转换为PIL格式图片
    draw = ImageDraw.Draw(img_pil)
    # 绘制文字信息
    draw.text((1400, 300), text, font=font, fill=(0, 0, 0))                                         #将文本输入到图片上
    bk_img = np.array(img_pil)
    cv2.imwrite(save_path, bk_img)


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    driver = Edge(executable_path=r'D:\Drives\edgedriver_win64\MicrosoftWebDriver.exe');
    driver.get('http://111.164.113.185:8090/godTotalGoods/web_searchGoods.do');
    driver.maximize_window();

    #从excel中读入商品列表

    search_shot(driver,'木材','./inital_img/pictures.png')

    # 图片加入指定文字
    fontpath = './Font/1647227196547932.ttc'                                                        # 设置需要显示的字体
    font = ImageFont.truetype(fontpath, 45)
    Add_Str('./inital_img/pictures.png','./maked_img/pictures.png',font,"查无此商品")

    # 将更改的图片加入word




