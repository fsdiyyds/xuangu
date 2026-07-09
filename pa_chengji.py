import selenium
from selenium import webdriver
import pathlib
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import requests
import xlrd
from xlutils.copy import copy
import os
import re

# 基本信息
# 视频存放路径
options = webdriver.ChromeOptions()
# options.add_experimental_option("debuggerAddress", "127.0.0.1:5003")
# options.add_argument('--headless')
driver = webdriver.Edge()

path = img_dir = os.path.join(os.curdir, 'xueyuan_images')  # xueyuan_images是文件夹名称
if not os.path.isdir(path):
    os.mkdir(path)



def get_piccj(name, id):
    '''
     作用：发布好看视频
    '''

    # 进入创作者页面，并上传视频
    driver.get("http://ztjk.ceht.net/login!logKSByout.do")
    time.sleep(0.5)
    driver.maximize_window()
    driver.find_element(by=By.XPATH, value='//input[@name="applicant.name"]').send_keys(name)
    time.sleep(0.2)
    driver.find_element(by=By.XPATH, value='//input[@name="applicant.idCardNo"]').send_keys(id)
    time.sleep(0.2)
    driver.find_element(by=By.XPATH, value='//button[@id="loginStu"]').click()
    time.sleep(0.5)

    driver.find_element(by=By.XPATH, value='//li[@id="cjcx"]').click()
    time.sleep(0.5)

    pic = driver.find_element(by=By.XPATH, value='//div[@class="flex-list-no-data2"]/img').get_attribute('src')
    print(pic)
    time.sleep(0.5)

    r = requests.get(pic)
    r.raise_for_status()
    pic_name = name + '.jpg'
    path_pic = os.path.join(img_dir, pic_name)

    with open(path_pic, 'wb') as f:
        f.write(r.content)
        f.close()
        print('保存成功')

    cj = driver.find_elements(by=By.XPATH, value='//div[@class="declaration_r"]')
    if len(cj) > 1:
        text = cj[1].text
    else:
        text = cj[0].text

    print(text)
    time.sleep(0.2)
    return text

def xieru():
    r_xls = xlrd.open_workbook("建筑信息模型技术员人员名册.xls")  # 读取excel文件
    sheet1_object = r_xls.sheets()[0] # 获取已有的行数
    excel = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
    worksheet = excel.get_sheet(0)  # 获取要操作的sheet
    rows_generator = list(sheet1_object.get_rows())
    index = 2
    for rows in rows_generator[2:]:
        print(rows)
        name, id = rows[1].value, rows[2].value
        print(name,type(id))
        print("{}的信息：{}".format(name, id))
        text_return = get_piccj(name, id)
        try:
            a = list(map(int, text_return))  # 可以替代上面两行代码
            print(a)
            worksheet.write(index, 8, a[0])
            worksheet.write(index, 9, a[1])
        except:
            worksheet.write(index, 8, '缺考')
            worksheet.write(index, 9, '缺考')
        worksheet.write(index, 10, text_return)
        index += 1
    excel.save("excelTest_real.xls")  # 保存并覆盖文件

xieru()


