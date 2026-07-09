# -*- coding: utf-8 -*-

'''
@Time    : 2021/12/6 15:02
@Author  : fenglei
@FileName: pashipin.py
@Software: PyCharm
 
'''
import os
import requests

url = 'https://haokan.baidu.com/videoui/api/videorec?tab=gaoxiao&act=pcFeed&pd=pc&num=20&shuaxin_id=1612592171486'
headers = {
  'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'
}
response = requests.get(url=url, headers=headers)
json_data = response.json()
videos = json_data['data']['response']['videos']
for index in videos:
  title = index['title']
  play_url = index['play_url']
  video_content = requests.get(url=play_url, headers=headers).content
  path = 'video\\'
  if not os.path.exists(path):
    os.mkdir(path)
  with open(path + title + '.mp4', mode='wb') as f:
    f.write(video_content)
    print('正在保存：', title)