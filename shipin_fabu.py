import selenium
from selenium import webdriver
import pathlib
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import glob
import re



options = webdriver.ChromeOptions()

options.add_argument('--user-data-dir=C:/Users/fenglei\AppData\Local\Google\Chrome/User Data/') #加载前面获取的 个人资料路径
# C:\Users\fenglei\AppData\Local\Microsoft\Edge\User
# options.add_experimental_option("debuggerAddress", "127.0.0.1:5003")
# , executable_path="/opt/google/chrome/chromedriver"

def publish_shipinhao(path_mp4,describe, short_biaoti):
    '''
     作用：发布微信视频号
    '''

    driver = webdriver.Chrome(options=options)

    # 进入微信视频号创作者页面，并上传视频
    driver.get("https://channels.weixin.qq.com/post/create")
    time.sleep(10)
    driver.find_element(by=By.XPATH, value='//input[@type="file"]').send_keys(path_mp4)

    # 等待视频上传完成
    # 检查一：等待正在处理文件的提示显示
    while True:
        time.sleep(3)
        try:
            driver.find_element(by=By.XPATH, value='//*[text()="取消上传"]')
            print("视频还在上传中···")
        except Exception as e:
            time.sleep(3)
            break;



    print("视频已上传完成！")

    # 输入视频描述
    driver.find_element(by=By.XPATH, value='//*[@data-placeholder="添加描述"]').send_keys(describe)
    time.sleep(1)

    # 添加位置
    driver.find_element(by=By.XPATH, value='//*[@class="position-display-wrap"]').click()
    time.sleep(2)
    driver.find_element(by=By.XPATH, value='//*[text()="不显示位置"]').click()
    time.sleep(2)

    driver.find_element(by=By.XPATH, value='//*[text()="视频为原创"]').click()

    driver.find_element(by=By.XPATH, value='//*[@placeholder="概括视频主要内容，字数建议6-16个字符"]').send_keys(short_biaoti)
    time.sleep(1)




    # 人工进行检查并发布
    time.sleep(3)
    # # 点击发布
    driver.find_element(by=By.XPATH, value='//*[text()="发表"]').click()
    time.sleep(10)


# 开始执行视频发布

for file_name in glob.glob('G:\剪\削球发布/*.mp4'):

    if len(file_name.split('\\')[-1].split('.mp4')) > 15:
        short_biaoti = re.sub('[^\u4e00-\u9fa5]+', '', file_name.split('\\')[-1].split('.mp4')[0][:15])
    else:
        short_biaoti = re.sub('[^\u4e00-\u9fa5]+', '', file_name.split('\\')[-1].split('.mp4')[0])

    describe ='#羽毛球#羽毛球大神#羽毛球双打 ' + file_name.split('\\')[-1].split('.mp4')[0]
    print(file_name, describe, short_biaoti)

    publish_shipinhao(file_name, describe, short_biaoti)