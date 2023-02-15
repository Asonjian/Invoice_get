from selenium.webdriver import Edge
import cv2
from docx import Document
from docx.shared import Cm
from PIL import ImageFont, ImageDraw, Image
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy

def rotateAntiClockWise90(img):  # 逆时针旋转90度
    trans_img = cv2.transpose(img)
    img90 = cv2.flip(trans_img, 0)
    return img90


def search_shot(driver, search_comment, save_path):
    """
    输入搜索内容后进行搜索并截图保存在相应目录
    @param driver: selenium控制变量
    @param search_comment: 将要搜索的内容
    @param save_path: 图片截图保存路径
    @return: none
    """
    try:
        driver.find_element("xpath", '/html/body/div[2]/div[2]/input').clear()
        driver.find_element("xpath", '/html/body/div[2]/div[2]/input').send_keys(search_comment)  # 定位输入文本框并将搜索内容输入
        driver.find_element("xpath", '/html/body/div[2]/div[2]/div').click()  # 定位按钮并点击搜索按键
        driver.get_screenshot_as_file(save_path)  # 保存图片
        return 1
    except():
        return 0


def Add_Str(open_path, save_path, font, text):
    """
    在图片中加入规定文字后保存图片
    @param open_path: 打开图片的路径
    @param save_path: 保存图片的路径
    @param font: 字体
    @param text: 插入的文本
    @return: None
    """
    img = cv2.imread(open_path)

    img_pil = Image.fromarray(img)  # cv2格式图片转换为PIL格式图片
    draw = ImageDraw.Draw(img_pil)
    # 绘制文字信息
    draw.text((1400, 300), text, font=font, fill=(0, 0, 0))  # 将文本输入到图片上
    bk_img = np.array(img_pil)
    img90 = rotateAntiClockWise90(bk_img)
    cv2.imwrite(save_path, img90)


def put_img_word(img_path, document, height, width):
    """
    将图片置入word中
    @param img_path: 图片路径
    @param document: 通过word打开的document对象

    @param height: 图片在word中显示的高
    @param width: 图片在word中显示的宽
    @return:document 返回更改的document对象
    """
    document.add_picture(img_path, width=Cm(width), height=Cm(height))
    return document


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':

    driver = Edge(executable_path=r'D:\Drives\edgedriver_win64\MicrosoftWebDriver.exe')
    driver.maximize_window()
    driver.get('http://111.164.113.185:8090/godTotalGoods/web_searchGoods.do')

    # 创建word-document对象
    document = Document()

    # 从excel中读入商品列表
    excel_in = xlrd.open_workbook('test.xlsx')
    sh_in = excel_in.sheet_by_index(0)
    count = sh_in.nrows

    # 新建excel对象进行excel的备注修改
    excel_out = copy(excel_in)
    sh_out = excel_out.get_sheet(0)
    font = xlwt.Font()                          # 创建字体对象
    style1 = xlwt.XFStyle()                     # 创建样式对象
    style1.font = font

    for i in range(0, count):
        goods_name = sh_in.cell_value(i, 1)  # 商品名称为表中第i行第2列中的值
        # 搜索并截图
        initial_path='./initial_img/'+str(i)+'.png'                                      # 原始截图保存路径
        prime_path='./maked_img/'+str(i)+'.png'                                        # 修改后的截图保存路径
        flag = search_shot(driver, goods_name, initial_path)
        if flag == 0:
            driver.get('http://111.164.113.185:8090/godTotalGoods/web_searchGoods.do')          # 失败后重新加载页面防止页面html变化失效

            sh_out.write(i, 4, '搜索涉及敏感词汇，截图有误！', style1)                                # 将失败备注写入excel
        else:
            # 图片加入指定文字
            fontpath = './Font/1647227196547932.ttc'  # 设置需要显示的字体
            font = ImageFont.truetype(fontpath, 45)
            Add_Str(initial_path, prime_path, font, "查无此商品")

            # 将更改的图片加入word
            document = put_img_word(prime_path, document, 23.55, 15.19)
            # 将成功备注写入excel
            sh_out.write(i, 4, '操作成功', style1)
    # word和excel修改后的保存
    document.save('demo.docx')
    excel_out.save('test.xlsx')
