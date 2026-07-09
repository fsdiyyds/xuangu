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
catalog_mp4 = r"G:/剪/视频发布测试"
# 视频描述
describe = "裸眼3D看蜘蛛侠 #搞笑 #电影 #视觉震撼"
time.sleep(1)
options = webdriver.EdgeOptions()
# options.add_experimental_option("debuggerAddress", "127.0.0.1:5003")
# d_ticket=2cd393a786e9d01e182624eaba9ee809a4a9a; n_mh=VNvU2J8XuveZuIwaX7BsxyIXQoWlHAs6tggTS8F-lrI; sso_uid_tt=f6bf5aa5d86b86bf783df92573b0a924; sso_uid_tt_ss=f6bf5aa5d86b86bf783df92573b0a924; toutiao_sso_user=15835580ea05ba407e24c7e40ac50788; toutiao_sso_user_ss=15835580ea05ba407e24c7e40ac50788; sid_ucp_sso_v1=1.0.0-KGExZGYyNWI5YjZhNTIxZTA1NWRmZjgyMzc3OGUzYjFhYzhjODgzZjgKHAjll9LM-QEQ8JaxpgYYGCAMMJXq7swFOAJA8QcaAmxmIiAxNTgzNTU4MGVhMDViYTQwN2UyNGM3ZTQwYWM1MDc4OA; ssid_ucp_sso_v1=1.0.0-KGExZGYyNWI5YjZhNTIxZTA1NWRmZjgyMzc3OGUzYjFhYzhjODgzZjgKHAjll9LM-QEQ8JaxpgYYGCAMMJXq7swFOAJA8QcaAmxmIiAxNTgzNTU4MGVhMDViYTQwN2UyNGM3ZTQwYWM1MDc4OA; sid_guard=8bb1497d94cddf6cf44638cd72d77d12%7C1691110257%7C5184001%7CTue%2C+03-Oct-2023+00%3A50%3A58+GMT; uid_tt=b7d455c44d6860af92235cb940e81880; uid_tt_ss=b7d455c44d6860af92235cb940e81880; sid_tt=8bb1497d94cddf6cf44638cd72d77d12; sessionid=8bb1497d94cddf6cf44638cd72d77d12; sessionid_ss=8bb1497d94cddf6cf44638cd72d77d12; sid_ucp_v1=1.0.0-KDIwODEwYmI0NTE1MmEyNGU3ZmNhMmE3ZWExNGZlMTEwMWNiZDYyNGEKFgjll9LM-QEQ8ZaxpgYYGCAMOAJA8QcaAmxmIiA4YmIxNDk3ZDk0Y2RkZjZjZjQ0NjM4Y2Q3MmQ3N2QxMg; ssid_ucp_v1=1.0.0-KDIwODEwYmI0NTE1MmEyNGU3ZmNhMmE3ZWExNGZlMTEwMWNiZDYyNGEKFgjll9LM-QEQ8ZaxpgYYGCAMOAJA8QcaAmxmIiA4YmIxNDk3ZDk0Y2RkZjZjZjQ0NjM4Y2Q3MmQ3N2QxMg; store-region=cn-sn; store-region-src=uid; odin_tt=eeabed34e713b46795e53ac03ed077ed721def21909a6dcaa837ea49f8619fffd1ccb7a49e26d1b7a6c04545f9312da1; __feed_out_channel_key=sports; _S_WIN_WH=1865_937; _S_DPR=1; _S_IPAD=0; _tea_utm_cache_24={%22utm_medium%22:%22wap_search%22}; notRedShot=1; _ga_1Y7TBPV8DE=GS1.1.1691747525.2.0.1691747527.0.0.0; _ga=GA1.1.1802202013.1690255443; local_city_cache=%E8%A5%BF%E5%AE%89; tt_scid=gPYa5aavHQyYttwmGP65bAMDttfo10yAR8VTVwdmwjVG6VbIIeBAAiZbvdVUfzs21e60; _ga_QEHZPBE5HH=GS1.1.1694078625.14.0.1694078625.0.0.0; tt_anti_token=LokXTSeuask-4bf30ebe9d90338a3ba87328e4f93ffb27485ec6f2075b1b53fba562158adad9; ttwid=1%7Cf4pTwS7z0rf8Tr-kwM4T1q6IIrg7diIutqgUFj7DbqI%7C1694078873%7C3d0b2e054cd6caad3dce934ce9cd15d8f201595387c173efe0d6abc2ec3b7419; xg_p_tos_token=422e927a54339a12f3baa3137f11dacd

driver = webdriver.Edge()


driver.get("https://www.toutiao.com")
driver.maximize_window()
time.sleep(30)
cookies = driver.get_cookies()
print(cookies)
#
coo = [{'domain': '.ctrip.com', 'httpOnly': False, 'name': '_bfi', 'path': '/', 'secure': False, 'value': 'p1%3D100021%26p2%3D10320670296%26v1%3D2%26v2%3D1'}, {'domain': '.ctrip.com', 'expiry': 1621428453, 'httpOnly': False, 'name': '_bfs', 'path': '/', 'secure': False, 'value': '1.2'}, {'domain': '.ctrip.com', 'expiry': 1684498653, 'httpOnly': False, 'name': '_bfa', 'path': '/', 'secure': False, 'value': '1.1621426633310.2sev1k.1.1621426633310.1621426633310.1.2'}, {'domain': '.ctrip.com', 'expiry': 1652962652, 'httpOnly': False, 'name': 'IsPersonalizedLogin', 'path': '/', 'secure': False, 'value': 'T'}, {'domain': '.ctrip.com', 'expiry': 1624018651, 'httpOnly': True, 'name': 'ticket_ctrip', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'bJ9RlCHVwlu1ZjyusRi+ypZ7X2r4+yoj3FJp0szEIZ+4fCVfpIpE2Ih8FO45kaoBwwWXoMI2Qo4Ae62WXZ7PkRcHrXbQLj5/9FBZN/p014UvPkF6fxTH5c8/gYL/XSwoG4l9Z6KSzVs+D8GPHliWO6xf2IvaPZHvMkf/yWcbNhFT+es0GveXrL7smbVgBDI/8YOBuq6RB1N9WgS/xvsxSlXvXXNffBte1JFsyxgRl+3TNRcxdoiYd7Zbrd8bSN922U+M3wtEhWV6l68F4kARSHmSpcQDjJhXjlW7VTNmknE='}, {'domain': '.ctrip.com', 'expiry': 3198226652, 'httpOnly': False, 'name': 'AHeadUserInfo', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'VipGrade=10&VipGradeName=%BB%C6%BD%F0%B9%F3%B1%F6&UserName=&NoReadMessageCount=1'}, {'domain': '.ctrip.com', 'expiry': 3198226652, 'httpOnly': False, 'name': 'login_type', 'path': '/', 'secure': False, 'value': '0'}, {'domain': '.ctrip.com', 'expiry': 1624018651, 'httpOnly': False, 'name': 'DUID', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'u=AB499F898293D4213913F0904DA89A5C&v=0'}, {'domain': '.ctrip.com', 'expiry': 1624018651, 'httpOnly': False, 'name': 'IsNonUser', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'F'}, {'domain': '.ctrip.com', 'expiry': 4042022400, 'httpOnly': False, 'name': '_RDG', 'path': '/', 'secure': False, 'value': '28c12c7bc6775b25342da625036ef08da9'}, {'domain': '.ctrip.com', 'expiry': 3198226652, 'httpOnly': False, 'name': 'login_uid', 'path': '/', 'secure': False, 'value': '5CD6CA24E085D8AFB663D43B371417C4'}, {'domain': '.ctrip.com', 'expiry': 4042022400, 'httpOnly': False, 'name': '_RGUID', 'path': '/', 'secure': False, 'value': 'ae6cce13-839e-4f11-bdda-c5b47717438e'}, {'domain': '.ctrip.com', 'expiry': 1652962652, 'httpOnly': False, 'name': 'UUID', 'path': '/', 'secure': False, 'value': '9EF7B9ABE7AC44929F7A401CD31AA30E'}, {'domain': '.ctrip.com', 'expiry': 1624018651, 'httpOnly': True, 'name': 'cticket', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': '582C72022D39A86F32229B0967C747E7AA6CA7E05C4970E05FCC097B253C7D09'}, {'domain': '.ctrip.com', 'expiry': 4042022400, 'httpOnly': False, 'name': '_RSG', 'path': '/', 'secure': False, 'value': 'Hyad3EXpivAEEq5tHRnQtB'}, {'domain': '.ctrip.com', 'expiry': 4042022400, 'httpOnly': False, 'name': '_RF1', 'path': '/', 'secure': False, 'value': '183.193.169.248'}]
for cookie in coo:
    driver.add_cookie(cookie)
time.sleep(2)
driver.get("https://passport.ctrip.com/user/login")
driver.find_element_by_id('personpwd').send_keys('输入自己的密码')
driver.find_element_by_id('personSubmit').click()
time.sleep(5)
driver.quit()





path = pathlib.Path(catalog_mp4)

# 视频地址获取
path_mp4 = ""
for i in path.iterdir():
    if(".mp4" in str(i)):
        path_mp4 = str(i);
        break;

if(path_mp4 != ""):
    print("检查到视频路径：" + path_mp4)
else:
    print("未检查到视频路径，程序终止！")
    exit()

# 封面地址获取
path_cover = ""
for i in path.iterdir():
    if(".png" in str(i) or ".jpg" in str(i)):
        path_cover = str(i);
        break;

if(path_cover != ""):
    print("检查到封面路径：" + path_cover)
else:
    print("未检查到封面路径，程序终止！")


def publish_haokan():
    '''
     作用：发布好看视频
    '''

    # 进入创作者页面，并上传视频
    # 进入创作者页面，并上传视频
    driver.get("https://mp.toutiao.com/profile_v4/xigua/upload-video?from=toutiao_pc")
    time.sleep(0.5)
    driver.maximize_window()

    driver.find_element(by=By.XPATH, value='//input[@name="applicant.name"]').send_keys(path_mp4)

    # 等待视频上传完成
    while True:
        time.sleep(3)
        try:
            driver.find_element_by_xpath('//*[text()="上传成功"]')
            break;
        except Exception as e:
            print("视频还在上传中···")

    print("视频已上传完成！")

    # 选择分类
    driver.find_element_by_xpath('//*[@class="hk-select-selection-item"]').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@title="影视"]').click()
    time.sleep(1)

    # 添加封面
    driver.find_element_by_xpath('//*[text()="上传封面"]').click()
    time.sleep(5)
    driver.find_element_by_xpath('//*[text()="本地上传"]').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@class="image-uploader-container"]//input[@type="file"]').send_keys(path_cover)
    time.sleep(5)
    driver.find_element_by_xpath('//*[text()="完成"]').click()

    time.sleep(2)
    # 输入标题
    driver.find_element_by_xpath('//*[@class="input-content"]//input').send_keys(Keys.CONTROL, 'a')
    time.sleep(2)
    driver.find_element_by_xpath('//*[@class="input-content"]//input').send_keys(describe)

    # 输入视频描述
    driver.find_element_by_xpath('//textarea').send_keys(describe)

    # 位置
    time.sleep(1)
    driver.find_element_by_xpath('//*[@class="location-input-wrap"]//input').send_keys("中关村人工智能科技园")
    time.sleep(3)
    driver.find_element_by_xpath('//*[text()="中关村人工智能科技园"]').click()

    # 参加话题
    time.sleep(1)
    driver.find_element_by_xpath('//*[text()="展开"]').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[text()="好看创作季"]').click()
    time.sleep(1)

    # 人工进行检查并发布
    # time.sleep(3)
    # # 点击发布
    # driver.find_element_by_xpath('//*[text()="发布"]').click()

# 开始执行视频发布
publish_haokan()
