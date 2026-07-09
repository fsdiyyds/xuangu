import selenium
from selenium import webdriver
import pathlib
import time
import glob
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

# 基本信息
# 视频存放路径
catalog_mp4 = r"G:\剪\削球发布"
# 视频描述

tag = ['#裸眼3D看蜘蛛侠', '#搞笑', '#电影', '#视觉震撼']
describe = "裸眼3D看蜘蛛侠 #搞笑 #电影 #视觉震撼"

options = webdriver.ChromeOptions()
options.add_argument('--user-data-dir=C:/Users/fenglei\AppData\Local\Google\Chrome/User Data/')  # 加载前面获取的 个人资料路径



def publish_toutiao(filename, title, tag):
    '''
     作用：发布趣头条视频
    '''
    path_cover = filename.split('.mp4')[0] + '.jpg'

    # 进入创作者页面，并上传视频
    driver = webdriver.Chrome(options=options)
    driver.get("https://mp.toutiao.com/profile_v4/xigua/upload-video?from=toutiao_pc")
    time.sleep(2)
    driver.find_element(by=By.XPATH, value='//input[@type="file"]').send_keys(filename)

    # 等待视频上传完成
    while True:
        time.sleep(3)
        try:
            driver.find_element(by=By.XPATH, value='//*[contains(text(),"上传成功")]')
            break;
        except Exception as e:
            print("视频还在上传中···")

    print("视频已上传完成！")
    placeholder = "请输入 1～30 个字符"

    # 输入标题
    # time.sleep(2)
    # driver.find_element(by=By.XPATH, value='//*[@placeholder="请输入 1～30 个字符"]').clear()
    time.sleep(2)
    driver.find_element(by=By.XPATH, value='//*[@placeholder="请输入 1～30 个字符"]').send_keys(title)


    # # 输入简介信息
    # time.sleep(1)
    # driver.find_element(by=By.XPATH, value='//textarea').clear()
    # time.sleep(2)
    # driver.find_element(by=By.XPATH, value='//textarea').send_keys(describe)

    # 添加封面
    time.sleep(1)
    driver.find_element(by=By.XPATH, value='//*[text()="上传封面"]').click()
    time.sleep(1)
    driver.find_element(by=By.XPATH, value='//*[text()="本地上传"]').click()
    time.sleep(1)
    # 准备选封面
    driver.find_element(by=By.XPATH, value='//*[@class=""detail""]/*/input[@type="file"]').send_keys(path_cover)
    time.sleep(3)
    driver.find_element(by=By.XPATH, value='//*[text()="确 定"]').click()

    # 声明原创
    time.sleep(1)
    driver.find_element(by=By.XPATH, value='//*[text()="原创"]').click()

    # # 选择分类
    # time.sleep(1)
    # driver.find_element(by=By.XPATH, value='//*[@placeholder="请选择分类"]').click()
    # time.sleep(1)
    # driver.find_element(by=By.XPATH, value='//*[text()="电影"]').click()
    # time.sleep(1)
    # driver.find_element(by=By.XPATH, value='//*[text()="分类："]').click()
    # time.sleep(1)

    # 输入标签


    time.sleep(1)
    for i in tag:
        driver.find_element(by=By.XPATH, value='//*[@class="content-tag"]//input').click()
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='//*[@class="content-tag"]//input').send_keys(i)
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='//*[@class="content-tag"]//input').send_keys(Keys.ENTER)
        time.sleep(1)



    # 人工进行检查并发布
    # time.sleep(3)
    # # 点击发布
    # driver.find_element_by_xpath('//*[text()="发布"]').click()


# 开始执行视频发布
files_mp4 = catalog_mp4 + '/*.mp4'
print(files_mp4)

for file_name in glob.glob(files_mp4):
    title = file_name.split('\\')[-1].split('.mp4')[0]
    print(file_name, title)

    publish_toutiao(file_name, title, tag)
