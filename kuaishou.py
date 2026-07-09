from os.path import join
import time
import requests
import json
import re
dir_name = 'G:\\2023年工作\\新建文件夹'
#响应头，整个复制之后，利用ctrl+r勾选正则表达式来替换（上面原来的(.*?): (.*) （冒号后面的空格）下面替换的格式 ‘$1': '$2',(冒号后面的空格，最后加逗号分隔）
headers = {
    'accept':  '*/*',
    'Accept-Encoding':  'gzip, deflate, br',
    'Accept-Language':  'zh-CN,zh;q=0.9',
    'Connection':  'keep-alive',
    'Content-Length':  '1380',
    'content-type':  'application/json',
    'Cookie':  'kpf=PC_WEB; clientid=3; did=web_a3ffd37c01f559eddb550ee7e087a3a8; userId=1225163093; kpn=KUAISHOU_VISION; kuaishou.server.web_st=ChZrdWFpc2hvdS5zZXJ2ZXIud2ViLnN0EqABuIfQIe7kp9Q6VBpDszIw1c0Yg-PLNNF2yhSGBMPVC26mRcyh07LRYOPWQxqnDCW_fyRTuCyohOV4nNQ0HE_2fUl6eqZEkGr9AQJET3jr9P0Gumg9C2cW7JqvPf8E7Q4SJ4m3AbZEmQd3xjoUFeBS-U30y34E_ZvHRjvI74FwOh8MbbozkmSYHARUQumH1Mk5RFP4WoZTAgmEMVQp_jYinxoSlCobbmtjkvjpY9x730BPP_C5IiBX1dFQ9or0yj_MSmTUVDSyIyDK3w4jHufneP1lKs3sfygFMAE; kuaishou.server.web_ph=bfeeee4b2a36e1a74bd13590f0691d21a3da'
   ,
    'Host':  'www.kuaishou.com',
    'Origin':  'https://www.kuaishou.com',
    'Referer':  'https://www.kuaishou.com/profile/3xq5w9ss8drgf3g',
    'sec-ch-ua':  '" Not A;Brand";v="99", "Chromium";v="96", "Google Chrome";v="96"',
    'sec-ch-ua-mobile':  '?0',
    'sec-ch-ua-platform':  '"Windows"',
    'Sec-Fetch-Dest':  'empty',
    'Sec-Fetch-Mode':  'cors',
    'Sec-Fetch-Site':  'same-origin',
    'User-Agent':  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36 Edg/117.0.2045.31',
}


def get_video(pcursor):
    baseurl = "https://www.kuaishou.com/graphql"
    data = {
        "operationName": "visionProfilePhotoList",
        "variables": {
            "userId": "3x8axahztu8bsdk",
            "pcursor": pcursor,
            "page": "profile"
        },
        "query": "fragment photoContent on PhotoEntity {\n  id\n  duration\n  caption\n  originCaption\n  likeCount\n  viewCount\n  realLikeCount\n  coverUrl\n  photoUrl\n  photoH265Url\n  manifest\n  manifestH265\n  videoResource\n  coverUrls {\n    url\n    __typename\n  }\n  timestamp\n  expTag\n  animatedCoverUrl\n  distance\n  videoRatio\n  liked\n  stereoType\n  profileUserTopPhoto\n  musicBlocked\n  __typename\n}\n\nfragment feedContent on Feed {\n  type\n  author {\n    id\n    name\n    headerUrl\n    following\n    headerUrls {\n      url\n      __typename\n    }\n    __typename\n  }\n  photo {\n    ...photoContent\n    __typename\n  }\n  canAddComment\n  llsid\n  status\n  currentPcursor\n  tags {\n    type\n    name\n    __typename\n  }\n  __typename\n}\n\nquery visionProfilePhotoList($pcursor: String, $userId: String, $page: String, $webPageArea: String) {\n  visionProfilePhotoList(pcursor: $pcursor, userId: $userId, page: $page, webPageArea: $webPageArea) {\n    result\n    llsid\n    webPageArea\n    feeds {\n      ...feedContent\n      __typename\n    }\n    hostName\n    pcursor\n    __typename\n  }\n}\n"
    }  #"pcursor"控制翻页（手动在网页中下滑之后会出现两个数据包）【这里要是字符类型】
    #页面搜索视频名字，然后找到抓包，再找响应网址
    # baseurl = "https://www.kuaishou.com/graphql"
    #headers有一个  'content-type':  'application/json',  这个定义了data(这里类似账号密码之类的数据)，要求data是json字符串

    data = json.dumps(data)  #将data由字典类型转换为字符串类型
    time.sleep(2)
    #发送请求，url:链接地址，headers:伪装，data:查询参数
    request = requests.post(url=baseurl,headers=headers,data=data, timeout=10)
    response = request.json()
    # pprint.pprint(response)
    ##字典数据利用键来找值  {"键":"值"} |列表直接利用位置索引 [值][值]  [0][1]
    # print('#############', response['data']['visionProfilePhotoList']['feeds'])
    print('***********', response['data']['visionProfilePhotoList']['pcursor'])

    pcursor = response['data']['visionProfilePhotoList']['pcursor']
     # title_list = response['data']['visionSearchPhoto']['feeds'][5]['photo']['caption']
    # print(title_list)
    # url_list = response['data']['visionSearchPhoto']['feeds'][5]['photo']['photoUrl']
    # print(url_list)
    feeds_list = response['data']['visionProfilePhotoList']['feeds']

    # print(feeds_list)
    k = 0
    for feeds in feeds_list:
        k =  k +1
        #每个feeds是feeds_list列表当中的一个个字典
        # print(feeds)  #利用这条可以把每个视频的信息都分别打印出来
        title = feeds['photo']['caption']
        list = feeds['photo']['photoUrl']
    # #下面这个打印出来把所有类似的数据都放在了同一个列表当中，与下载无关
    # # titles = [i['photo']['caption']for i in feeds_list]
    # # print(titles)
    #     list = [i['photo']['photoUrl']for i in feeds_list]
    #     print('视频的地址为', list)
    #     time.sleep(2)
        ##保存视频  【搜索关键词下载视频/知道一个用户的视频/翻页下载】
        new_title = re.sub(r'[\/:*?"<>|\n]','_',title)  #在windows操作系统当中，必须是没有一些特殊字符  #标题过长可以替换（字符串的切片）当>=256


        # response1 = requests.get(list)
        #
        # expected_length = response1.headers.get('Content-Length')
        # if expected_length is not None:
        #     actual_length = response.raw.tell()
        #     expected_length = int(expected_length)
        #     if actual_length < expected_length:
        #         raise IOError(
        #             'incomplete read ({} bytes read, {} more expected)'.format(
        #                 actual_length,
        #                 expected_length - actual_length
        #             )
        #         )
        # 发送网络请求，请求每一个视频地址，获取视频二进制数据
        try:
            mp4_data = requests.get(list, timeout=20).content
            new_title = re.sub('[^\u4e00-\u9fa5]+', '', new_title)
            filename = new_title + '%s.mp4'
            with open(join(dir_name,filename),mode='wb') as f:
                f.write(mp4_data)
                print(new_title,"下载完成")
        except:
            continue

    return pcursor


pcursor = '1.62796326E12'
while pcursor != 'no_more':
    time.sleep(1)
    pcursor = get_video(pcursor)


    # mp4_data.close()
