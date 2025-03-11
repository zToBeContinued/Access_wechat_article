import os
import re  # 使用正则表达式
import json  # 用于json转码
import time
import random
import jsonpath
import requests
import pandas as pd  # 修改excel
from bs4 import BeautifulSoup
from fake_useragent import UserAgent  # 生成随机浏览器标识
import sys
import logging

logging.basicConfig(  # 配置日志记录
    filename='app.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
requests.packages.urllib3.disable_warnings()  # 去除网络请求警告


class AccessPosts:
    """功能：
        1.根据URL访问原页面获取网页文本数据
        2.按保存规则进行存储（单个、批量）
        3.检测人机验证，并去除验证
    """

    def __init__(self):
        self.root_path = r'./all_data/'  # 数据存储目录
        self.official_names_head = '公众号----'  # 公众号保存目录开头，用以保存对应公众号的信息，公众号: xxx
        self.headers = {
            'User-Agent': UserAgent().random,  # 生成随机的浏览器标识头
        }
        self.cookies = {"poc_sid": ''}  # 用以保存设备ID，用来去除人机验证
        os.makedirs(self.root_path, exist_ok=True)  # 创建保存路径，如果文件夹已存在，则忽略，默认为r'./all_data'

    def save_one_article(self, article_content, img_save_flag=True, content_save_flag=True):
        """输入：文章文本内容，是否保存图片（默认保存），是否保存文章内容到文件（默认保存）
           输出：保存flag
           功能：整理文本内容，创建保存路径
        """
        # 整理文章关键信息
        nickname = re.search(r'var nickname.*"(.*?)".*', article_content).group(1)  # 公众号名称
        article_link = re.search(r'var msg_link = .*"(.*?)".*', article_content).group(1)  # 文章链接
        createTime = re.search(r"var createTime = '(.*?)'.*", article_content).group(1)  # 文章创建时间
        # year, month, day = createTime.split(" ")[0].split("-")      # 年，月，日
        # hour, minute = createTime.split(" ")[1].split(":")          # 小时，分钟
        author = re.search(r'var author = "(.*?)".*', article_content).group(1)  # 文章作者
        article_title = re.search(r"var title = '(.*?)'.*", article_content).group(1)  # 文章标题
        article_title_win = re.sub(r'[\\/*?:"<>|].', '_', article_title)  # Windows下标题
        article_title_win = article_title_win.replace('.', '')  # Windows下标题，去除小数点，防止自动省略报错

        # 创建公众号保存目录
        official_path = self.root_path + self.official_names_head + nickname  # 各种公众号存储根路径
        os.makedirs(official_path, exist_ok=True)

        """下载文章图片"""
        if img_save_flag:  # 类属性中开启保存选项！
            print('开启保存文章图片选项，准备下载文章图片')
            # 创建文章图片保存目录
            img_save_path = (self.root_path + self.official_names_head + nickname + '/'  # 图片保存路径
                             + createTime.replace(':', '：') + ' ' + article_title_win)
            os.makedirs(img_save_path, exist_ok=True)  # 创建图片保存目录

            # 保存该文章图片内容
            images = article_content.split('https://mmbiz.qpic.cn/')
            # print(images)
            for i in range(0, len(images) - 1):
                image_url = 'https://mmbiz.qpic.cn/' + images[i + 1].split('"')[0]
                # print('正在获取图片：' + image_url)
                image_name = ''
                response = requests.get(image_url, cookies=self.cookies, verify=False)
                if response.status_code == 200:
                    # 图片命名
                    img_hz = ['gif', 'jpg', 'jpeg', 'png', 'webp']
                    for imghz in img_hz:
                        if imghz in image_url:
                            image_name = str(i + 1) + '.' + imghz
                    if image_name == '':  # 如果链接中没有标明图片属性
                        image_name = str(i + 1) + '.jpg'
                    file_path = img_save_path + '/' + image_name
                    # 保存图片
                    with open(file_path, 'wb') as f:
                        f.write(response.content)
                    print(f"已成功下载图片： {file_path}")
                else:
                    print(f"无法下载图片，状态码: {response.status_code}")
            print('已保存文章图片>>>> ' + article_title)

        """保存文章文本内容"""
        if content_save_flag:
            # 将文字内容转换为列表形式存储
            soup = BeautifulSoup(article_content, 'html.parser')
            original_texts = soup.getText().split('\n')  # 将页面所有的文本内容提取，并转为列表形式
            article_texts = list(filter(lambda x: bool(x.strip()), original_texts))  # filter() 函数可以根据指定的函数对可迭代对象进行过滤

            # 创建 or 打开表格，检查文件是否存在，判断不存在时创建表格文件
            article_contents_path = official_path + '/' + '文章内容(article_contents).xlsx'  # 文章内容文件路径
            if not os.path.exists(article_contents_path): pd.DataFrame().to_excel(article_contents_path, index=False)
            frame_df = pd.read_excel(article_contents_path)  # 读取表格内容，默认打开DataFrame对象包含第一个工作表中的数据

            # 将新数据转换为 DataFrame 并添加到现有 DataFrame 的末尾
            local_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 本地时间
            columns = ['本地存储时间', '文章发布时间', '文章名称', '文章链接', '文章文本内容']  # 列名
            new_data_df = pd.DataFrame([[local_time, createTime, article_title, str(article_link), str(article_texts)]],
                                       columns=columns)
            df = pd.concat([frame_df, new_data_df], ignore_index=True)

            # 将更新后的数据写入 Excel 文件
            df.to_excel(article_contents_path, index=False)
            print(local_time + ' 已保存文章>>>> ' + article_title)
            print(local_time + ' 内容存储路径>>>> ' + article_contents_path)

    def get_one_article(self, url, img_save_flag=True, content_save_flag=True):
        """
            输入：微信文章链接（永久链接或短链接），是否保存图片（默认保存），是否保存文章内容到文件（默认保存）
            输出：无（内容保存目录在终端显示）
        """
        res = requests.get(url, headers=self.headers, cookies=self.cookies, verify=False)  # 发起请求
        # 验证请求
        if 'var createTime = ' in res.text:  # 正常获取到文章内容
            print('正常获取到文章内容，开始保存操作')
            try:
                self.save_one_article(res.text, img_save_flag, content_save_flag)  # 开始保存单篇文章
                return {'content_flag': 1, 'content': res.text}  # 用来获取公众号主页链接
            except:
                article_title = re.search(r"var title = '(.*?)'.*", res.text)  # 文章标题
                if article_title: article_title = article_title.group(1)
                print('检测到抓取出错，文章名>>>>    ' + article_title)
                print('检测到抓取出错，文章链接>>>>    ' + url)
                return {'content_flag': 0}
        elif '>当前环境异常，完成验证后即可继续访问。<' in res.text:
            print('当前环境异常，请检查链接后访问！！！')  # 代码访问遇到人机验证，需进行验证操作
            return {'content_flag': 0}
        elif '操作频繁，请稍后再试。' in res.text:
            print('操作频繁了，等会再弄或换ip弄！！！')  # 遇到次数较少，如有遇到请前往GitHub留言
            return {'content_flag': 0}
        else:
            print('出现其他问题，请查找原因后再试！！！！')  # 出现错误信息，如有遇到请前往GitHub留言
            return {'content_flag': 0}

    def get_list_article(self, name_link, img_save_flag=True, content_save_flag=True):
        """ 输入：公众号名称或公众号的一篇文章，是否保存图片（默认保存），是否保存文章内容到文件（默认保存）
            输出：无（内容保存目录在终端显示）
            功能：保存文章列表中所有内容
        """
        if 'http' in name_link:
            print('检测到输入为链接，开始获取公众号名称')
            content = self.get_one_article(name_link, False, False)
            if content['content_flag'] == 1:
                nickname = re.search(r'var nickname.*"(.*?)".*', content['content']).group(1)  # 公众号名称
            else:
                print('未获取到公众号名称')
                return None
        else:
            print('检测到输入为公众号名称')
            nickname = name_link

        official_path = self.root_path + self.official_names_head + nickname  # 公众号存储根路径
        # article_contents_path = official_path + '/' + '文章列表（article_list）_原始链接.xlsx'  # 文章内容文件路径
        article_list_path = official_path + '/' + '文章列表（article_list）_直连链接.xlsx'  # 文章列表文件路径
        if not os.path.exists(article_list_path):  # 如果文件不存在
            print('文件不存在，请检查目录文件>>>>  文章列表（article_list）_直连链接.xlsx')
        else:
            frame_df = pd.read_excel(article_list_path)  # 读取表格内容，默认打开DataFrame对象包含第一个工作表中的数据
            # 开始下载文章内容
            for index, row in frame_df.iterrows():
                roll_url = row.iloc[4]  # 获取直连链接
                self.get_one_article(roll_url, img_save_flag, content_save_flag)

    # def verify_user(self, url, content):
    #     """
    #         输入：url=请求路径，content=网页内容，如：res.text，遇到此情况时使用：  >当前环境异常，完成验证后即可继续访问。<
    #         输出：验证标志（1为有效），网页内容，cookie值
    #             {'verify_flag': 1, 'content': res.text, 'poc_sid': poc_sid}
    #         poc_sid == deviceID
    #     """
    #     print('开始验证，正在获取参数poc_sid')
    #     poc_token = re.search(r'poc_token.*"(.*?)"', content).group(1)
    #     poc_sid = re.search(r'poc_sid.*"(.*?)"', content).group(1)  # poc_sid为cookie参数
    #     cap_appid = re.search(r'cap_appid.*"(.*?)"', content).group(1)
    #     cap_sid = re.search(r'cap_sid.*"(.*?)"', content).group(1)
    #     target_url = re.search(r'target_url.*"(.*?)"', content).group(1)
    #
    #     try:
    #         '''验证请求第一步'''
    #         verify1_url = ('https://t.captcha.qq.com/cap_union_prehandle?' + 'protocol=https&accver=1&showtype=popup&'
    #                        'ua=TW96aWxsYS81LjAgKFdpbmRvd3MgTlQgMTAuMDsgV2luNjQ7IHg2NCkgQXBwbGVXZWJLaXQvNTM3LjM2IChLSFRNTCwgbGlrZSBHZWNrbykgQ2hyb21lLzEyNy4wLjAuMCBTYWZhcmkvNTM3LjM2IEVkZy8xMjcuMC4wLjA%3D&'
    #                        'noheader=0&fb=1&aged=0&enableAged=0&enableDarkMode=1&grayscale=1&dyeid=0&clientype=2'
    #                        '&aid=' + cap_appid + '&deviceID=' + poc_sid + '&sid=' + cap_sid +
    #                        '&cap_cd=&uid=&lang=zh-cn&elder_captcha=0&js=%2Ftcaptcha-frame.8d77d8b0.js&login_appid='
    #                        '&entry_url=https%3A%2F%2Fmp.weixin.qq.com%2Fmp%2Fwappoc_appmsgcaptcha&wb=1&version=1.1.0'
    #                        '&subsid=1&callback=_aq_873604&sess=')
    #         header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36 Edg/134.0.0.0'}
    #         verify1 = requests.get(verify1_url, headers=header, verify=False)
    #         # sess，pow_answer 为第三步中的请求数据
    #         sess = re.search('sess":"(.*?)"', verify1.text).group(1)
    #         pow_answer = re.search('prefix":"(.*?)"', verify1.text).group(1)
    #
    #         '''验证请求第二步'''
    #         verify2_url = 'https://t.captcha.qq.com' + re.search('tdc_path":"(.*?)"', verify1.text).group(1)
    #         verify2 = requests.get(verify2_url, headers=header, verify=False)
    #         # eks 为第三步中的请求数据
    #         eks = re.search(r"='(.*?)'", verify2.text).group(1)
    #
    #         '''验证请求第三步'''
    #         verify3_url = 'https://t.captcha.qq.com/cap_union_new_verify'
    #         verify3_data = {
    #             'collect': 'F97A58z6EKA4CNUjzdxrYiPXOGxCX1E4UPbmPhuuy6vojKPeA0EUN5DJWtjE3y0eow298aaKR+wKb7f8wsB6K1uaS93BwGTk8a18UNChBgwMYPRdHERtNoHs66mCG3FRhfxgEi758hvugEzzsKyNStp8ChZa9NqJ0OEBsVqsaTAZoVzkIZ8KqUgoMUW9EhoXesF5tqB9arGi+ZkBPrw5w0HzVR8yx1ehQhjixIw5rjCXg98Z2Fq8P4knkq9epFQEgB6vpR7K8gZ0VhmRCLXNTM4FsHnMdHWBX7orOllWdusPAlCMnsXMj7ucO9aDyP1e2fYsJYwK9zeSi8zvQ4F/XP6a9NvOYY4dZR7HI2UaJwUG0xxPU14zymkk8CWHWG5i1kKYGUz6X/yISfEczLkCMHECgJDtMOJzb9WhkuyfD7PyvpL1rU1lgWApFJp3c46RvCTftmfhfu2IJMTZ5LwWtxJIX8zsUj42pWiWM7iiqSzoH9gBgLGyJSWKUXm4f4jIeMj4V8hECgrYT5E9Oz1zl3Yib74HV2R8NjM6e9VjI7fu3/GKVdkQP0CgnSbYJzvJpDsECdY1CSgwEtI2AaC8x2eECThJ2j/3X9pb4ypH6N6ZSDWD5I67rOUeHLi8L0NN1ISm/HiGD8mWDOGLyyFsEGuGGzuMqy+Fxehtr2uvyxRWtWadGhG34osn1aNKcJcMK4iSERJeZGBbpTQNaA626rxzjjxBEbuNyRXvSHHbB33WzGT/74wrkaTRpcpwo6IGd8Rw93kThxuEpb8SmFVDAcIexRBn/+AWPnpcbM1aS82k0aXKcKOiBqTRpcpwo6IG0wdvP4uHCcyi67Tdt1B0yTKhZOtZ+z2xwh7FEGf/4BYAswz/QxMnADq0ZSBjHYmU9mPAx29tLUZQG3YboJK5sX+m4Ga9XlnqGW7hZEE83B/5powR4JuSyIFd38tOJdxgJBM9r8WKWUcRsXosFi1NHy7U7yuIT+bQ4HsRrnHlh73wJO3rsC5ShXfIRH6l4Q/Zf5HY0ENCIULn//Cv2azN4xGZsuYK4mxG2jtSpNboVeiJruE9wXg798jpf7CJKPV+v/ffl9+s5AbTJ69l3LT3td/cgtlkCpKsxrRZbD7ZI53YESSMtjzw76PGoKGD+MFQRkucxPcXQBEBOtd9zxxMIIZiXFWsG62+HLeQVL86apMlSFzJ8zNsU5xeilajsVkeqEWZAmdfoskf5iXmDpEoxSabuF0xUPRdpNGlynCjogak0aXKcKOiBqTRpcpwo6IGJ5hRMBQehUwRZc+Z+lhesnYtjIlRt+75qu5cfqEIgUYtg2+DknP0YDRuxdVoPCN1',
    #             'tlg': 1312,
    #             'eks': eks,
    #             'sess': sess,
    #             'ans': '[{"elem_id":0,"type":"DynAnswerType_TIME","data":""}]',
    #             'deviceID': poc_sid,
    #             'pow_answer': pow_answer + '#104',
    #             'pow_calc_time': 1,
    #         }
    #         header = {
    #             'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36 Edg/134.0.0.0',
    #             'Accept': 'application / json, text / javascript, * / *; q = 0.01',
    #             'Accept - Encoding': 'gzip, deflate, br, zstd',
    #             'Accept-Language': 'zh-CN,zh;q=0.9',
    #         }
    #         verify3 = requests.post(verify3_url, headers=header, data=verify3_data, verify=False)
    #         # ticket，randstr为第四步请求参数
    #         ticket = json.loads(verify3.text)['ticket']
    #         randstr = json.loads(verify3.text)["randstr"]
    #         print(ticket)
    #         print(randstr)
    #         print(verify3.text)
    #
    #         '''验证请求第四步'''
    #         verify4_url = 'https://mp.weixin.qq.com/mp/wappoc_appmsgcaptcha?action=Check&x5=0&f=json'
    #         verify4_data = {
    #             'target_url': target_url,
    #             'poc_token': poc_token,
    #             'appid': cap_appid,
    #             'ticket': ticket,
    #             'randstr': randstr,
    #         }
    #         self.cookies['poc_sid'] = poc_sid  # 重置类属性 cooikes的值
    #         verify4 = requests.post(verify4_url, headers=header, cookies=self.cookies, data=verify4_data,
    #                                 verify=False)
    #         # print(verify4.text)
    #         # print('发送成功后，poc_sid就可以正常使用了')
    #
    #         '''验证请求第五步'''
    #         modify_url = url + '&poc_token=' + poc_token
    #         res = requests.get(modify_url, headers=self.headers, cookies=self.cookies, verify=False)  # 发起请求
    #         print('已完成验证请求，后续请求若仍存在异常，请检查！')
    #         # print(res.text)
    #         return {'verify_flag': 1, 'content': res.text}
    #     except:
    #         print('验证失败，请检查后再进行尝试')
    #         return {'verify_flag': 0}


class ArticleDetail(AccessPosts):
    def __init__(self):
        super().__init__()
        self.biz = None
        self.uin = None
        self.key = None
        self.pass_ticket = None
        self.text = 'website'  # 预留位

    def get_article_link(self, url):
        """
            输入：公众号下任意一篇已发布的文章 短链接！！
            功能：通过公众号内的文章获取到公众号的biz值，拼接出公众号主页链接
        """
        content = super().get_one_article(url, False, False)  # 获取网页文本内容
        if content['content_flag'] == 1:
            print('正在生成微信公众号主页链接……\n')
            self.biz = re.search('var biz = "(.*?);', content['content']).group(1).replace('" || "', '').replace('"',
                                                                                                                 '')
            names = re.search(r'var nickname.*"(.*?)".*', content['content']).group(1)  # 公众号名称
            main_url = ('https://mp.weixin.qq.com/mp/profile_ext?action=home&__biz=' + self.biz +
                        '&scene=124#wechat_redirect')
            print(names + '公众号主页链接为：' + main_url)
            print('将此链接 （￣︶￣）↗　 粘贴发送到 ‘微信PC端-文件传输助手’')
        else:
            print('未获取到文章内容，请检查链接是否正确')

    def access_origin_list(self, access_token, pages=None, save_list=True, transform_list=True):
        """ 输入：access_token(从fiddler获取的链接)，保存页数（默认全部），是否保存到文件（默认保存），是否转换链接（默认转换）
            输出：无（获取的文章列表将保存在本地目录下）
            功能：
                ① 请求得到文章信息（文章标题、文章链接、文章创建日期）
                ②以excel文件形式存储，文件名设置为对应公众号的名称
        """
        # 检验access_token是否合法
        self.biz = str(re.search('biz=(.*?)&', access_token).group(1))
        self.uin = str(re.search('uin=(.*?)&', access_token).group(1))
        self.key = str(re.search('key=(.*?)&', access_token).group(1))
        self.pass_ticket = str(re.search('pass_ticket=(.*?)&', access_token).group(1))
        if self.biz and self.uin and self.pass_ticket and self.key:
            print('参数齐全，开始获取文章信息，默认状态获取全部文章')
        else:
            print('\n※※※ 参数有误，请重新输入')
            return None

        '''获取文章列表，格式化内容为一个二维数组：all_list'''
        all_list = None  # 用来存储获取的文章列表
        # 遍历公众号下所有文章链接
        if not pages:
            page = 0
            passage_list = []
            print('开始获取公众号下所有的文章列表')
            while True:
                p_data = self.get_next_list(page)
                if p_data['m_flag'] == 1:
                    for i in p_data['passage_list']:
                        passage_list.append(i)
                else:
                    print('请求结束，文章列表获取完毕！')
                    break
                page = page + 1
                delay_time = random.uniform(1, 5)  # 延迟时间
                print('为预防被封禁,开始延时操作，延时时间：' + str(delay_time) + '秒')
                time.sleep(delay_time)  # 模拟手动操作，随机延时delay_time秒，预防被封禁
            all_list = passage_list
        # 获取公众号下指定页数的文章链接
        else:
            print('输入值为：' + str(pages) + '，开始获取前' + str(pages) + '页文章')
            passage_list = []
            for pages in range(pages):
                p_data = self.get_next_list(pages)
                if p_data['m_flag'] == 1:
                    for i in p_data['passage_list']:
                        passage_list.append(i)
                else:
                    print('请求结束，文章列表获取完毕！')
                    break
                delay_time = random.uniform(1, 5)  # 延迟时间
                print('为预防被封禁,开始延时操作，延时时间：' + str(delay_time) + '秒')
                time.sleep(delay_time)  # 模拟手动操作，随机延时1-5秒，预防被封禁
            all_list = passage_list
        print('********************共获取到 ' + str(len(all_list)) + ' 篇文章，开始保存文章，若为0篇请检查错误！！！')
        if not all_list: print('获取到文章列表为空，请注意检查！！！！')
        if not all_list: return None  # 如果获取为空

        '''保存文章列表到文件，保存目录'''
        nickname = ''  # 临时放置公众号名称
        if save_list:
            print('****************************************开始保存文章，若以上为 获取到0篇 请检查错误！！！')
            # 首先获取公众号名称
            # new_url = all_list[0][2] + '&pass_ticket=' + self.pass_ticket + '&uin=' + self.uin + '&key=' + self.key
            new_url = all_list[0][3].replace('amp;', '')
            res = requests.get(new_url, headers=self.headers, verify=False)  # 使用微信客户端的token跳过验证
            nickname = re.search(r'var nickname.*"(.*?)".*', res.text).group(1)  # 公众号名称

            # 创建公众号保存目录
            official_path = self.root_path + self.official_names_head + nickname  # 各种公众号存储根路径
            os.makedirs(official_path, exist_ok=True)

            # 创建 or 打开表格，检查文件是否存在，判断不存在时创建表格文件
            article_contents_path = official_path + '/' + '文章列表（article_list）_原始链接.xlsx'  # 文章内容文件路径
            if not os.path.exists(article_contents_path): pd.DataFrame().to_excel(article_contents_path, index=False)
            frame_df = pd.read_excel(article_contents_path)  # 读取表格内容，默认打开DataFrame对象包含第一个工作表中的数据

            # 将新数据转换为 DataFrame 并添加到现有 DataFrame 的末尾
            columns = ['本地保存时间', '文章发布时间', '文章名称', '文章原始链接（直接访问会提示验证）']  # 列名
            new_data_df = pd.DataFrame(all_list, columns=columns)
            df = pd.concat([frame_df, new_data_df], ignore_index=True)

            # 将更新后的数据写入 Excel 文件
            df.to_excel(article_contents_path, index=False)
            local_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 本地时间
            print(local_time + ' 已获取公众号文章目录>>>> ' + nickname)
            print(local_time + ' 存储路径>>>> ' + article_contents_path)

        """转换 文章原始链接 为可直接访问链接"""
        if transform_list:
            print("开始转换 " + nickname + ' 公众号的文章列表原始链接')
            # 检测公众号的存储目录
            official_path = self.root_path + self.official_names_head + nickname  # 公众号存储根路径
            article_contents_path = official_path + '/' + '文章列表（article_list）_原始链接.xlsx'  # 文章内容文件路径
            article_list_path = official_path + '/' + '文章列表（article_list）_直连链接.xlsx'  # 文章列表文件路径
            if not os.path.exists(article_contents_path):  # 如果文件不存在
                print('文件不存在，请检查目录文件>>>>  文章列表（article_list）_原始链接.xlsx')
            else:
                frame_df = pd.read_excel(article_contents_path)  # 读取表格内容，默认打开DataFrame对象包含第一个工作表中的数据
                new_links = []  # 转换后的新链接存储

                # 修改短链接 方法1：删除元素“amp;”
                for index, row in frame_df.iterrows():
                    new_url = row.iloc[3].replace('amp;', '')  # 获取原始链接，并对其进行转化
                    new_links.append(new_url)  # 添加转化后的链接到数组中

                # # 修改短链接 方法2：添加pass_ticket、uin、key三个参数实现访问（此为临时链接！！！）
                # for index, row in frame_df.iterrows():
                #     new_url = row.iloc[2].replace('amp;', '')  # 获取第 3 列的值
                #     # res = requests.get(new_url, verify=False)  # 使用微信客户端的token跳过验证
                #     # print(index)

                # 合并 转换后的链接 到 原数据表，列合并操作
                frame_df['可直接访问链接'] = new_links  # 把列表作为新列添加到 DataFrame

                # 将更新后的数据写入 Excel 文件
                frame_df.to_excel(article_list_path, index=False)
                local_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 本地时间
                print(local_time + ' 已转换公众号文章列表>>>> ' + nickname)
                print(local_time + ' 存储路径>>>> ' + article_list_path)
        return all_list  # 返回

    def get_next_list(self, page):
        # 从0开始计数，第 0 页相当于默认页数据
        pages = int(page) * 10
        print('正在获取第 ' + str(page + 1) + ' 页文章列表')
        url = ('https://mp.weixin.qq.com/mp/profile_ext?action=getmsg&__biz=' + self.biz + '&f=json&offset='
               + str(pages) + '&count=10&is_ok=1&scene=124&uin=' + self.uin + '&key=' + self.key + '&pass_ticket='
               + self.pass_ticket + '&wxtoken=&appmsg_token=&x5=0&f=json')
        try:
            res = requests.get(url=url, headers=self.headers, timeout=10, verify=False)
        except:
            print('失败！！！获取第 ' + str(page + 1) + ' 页文章列表失败！！！')
            print('请检查错误类型，详情记录在日志中')
            exc_type, exc_value, exc_traceback = sys.exc_info()  # 获取当前异常的信息
            logging.error(f'发生异常: {exc_type.__name__}: {exc_value}', exc_info=True)
            res = ArticleDetail()  # 保证返回值不会报错
        if 'app_msg_ext_info' in res.text:
            # 解码json数据
            get_page = json.loads(json.loads(res.text)['general_msg_list'])['list']
            ''' get_page[0]为
            {'comm_msg_info': {'id': 1000000107, 'type': 49, 'datetime': 1722467332, 'fakeid': '3910318108', 'status': 2, 'content': ''}, 'app_msg_ext_info': {'title': '国务院7月重要政策', 'digest': '', 'content': '', 'fileid': 100007840, 'content_url': 'http://mp.weixin.qq.com/s?__biz=MzkxMDMxODEwOA==&amp;mid=2247491511&amp;idx=1&amp;sn=a36291fdee52a0f53d145edec8058e04&amp;chksm=c0084d6abbcac962a50153c89fe9c19b6f8b1c5e5ac50b05adcb49bdfad8638522ab426c3f4b&amp;scene=27#wechat_redirect', 'source_url': '', 'cover': 'https://mmbiz.qpic.cn/mmbiz_jpg/JRAjbHqmggrlZibDMibLP4ryNqhYXgolJOdQj2P8t2QQFVicickzAo7Gv1SzazwJY6lDylcanx2ic60HDbMvK8OKQpg/0?wx_fmt=jpeg', 'subtype': 9, 'is_multi': 1, 'multi_app_msg_item_list': [{'title': '8月起，这些新规将影响你我生活！', 'digest': '', 'content': '', 'fileid': 0, 'content_url': 'http://mp.weixin.qq.com/s?__biz=MzkxMDMxODEwOA==&amp;mid=2247491511&amp;idx=2&amp;sn=b3f5b6bcf8727c8c90fce7e588e6e7da&amp;chksm=c0eb20c99ca2f90032a6234002ed2cc9c2c000f87cff34f4d8d763878c0bb5275800db876ca7&amp;scene=27#wechat_redirect', 'source_url': '', 'cover': 'https://mmbiz.qpic.cn/mmbiz_jpg/JRAjbHqmggrc08yJMZ6CQ3VL6VzmEIymSUyATlL6o3xaDJJ0D2CtpQg31Vy7jdCaic86zqkgJ9oAFGyia78ZOq7g/0?wx_fmt=jpeg', 'author': '', 'copyright_stat': 100, 'del_flag': 1, 'item_show_type': 0, 'audio_fileid': 0, 'duration': 0, 'play_url': '', 'malicious_title_reason_id': 0, 'malicious_content_type': 0}, {'title': '8月，你好！', 'digest': '', 'content': '', 'fileid': 100007860, 'content_url': 'http://mp.weixin.qq.com/s?__biz=MzkxMDMxODEwOA==&amp;mid=2247491511&amp;idx=3&amp;sn=cd25de57b74b63b0f3b1a9888b9cd94d&amp;chksm=c0c7f30fdd5fc0ea4a2765f5fd29e1faeb0e352e888ee8556521ab23bc9528d68f42deaa9d15&amp;scene=27#wechat_redirect', 'source_url': '', 'cover': 'https://mmbiz.qpic.cn/mmbiz_jpg/JRAjbHqmggrlZibDMibLP4ryNqhYXgolJO9CnECAnMLDPY39Y9iarcFtM1ibrBvhKcGFyl1wicHysvTrYx4GfLybt8g/0?wx_fmt=jpeg', 'author': '', 'copyright_stat': 100, 'del_flag': 1, 'item_show_type': 0, 'audio_fileid': 0, 'duration': 0, 'play_url': '', 'malicious_title_reason_id': 0, 'malicious_content_type': 0}], 'author': '', 'copyright_stat': 100, 'duration': 0, 'del_flag': 1, 'item_show_type': 0, 'audio_fileid': 0, 'play_url': '', 'malicious_title_reason_id': 0, 'malicious_content_type': 0}}
            存储形式为二维数组，[[时间，文章标题，文章链接],[时间，文章标题，文章链接]
            '''
            passage_list = []  # 存放一页内的所有文章
            for i in get_page:
                # 时间戳转换
                time_tuple = time.localtime(i['comm_msg_info']['datetime'])
                create_time = time.strftime("%Y-%m-%d", time_tuple)
                title = i['app_msg_ext_info']['title']
                content_url = i['app_msg_ext_info']['content_url'].replace('#wechat_redirect', '')
                local_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 本地时间
                passage_list.append([local_time, create_time, title, content_url])
                if i['app_msg_ext_info']['multi_app_msg_item_list']:
                    for j in i['app_msg_ext_info']['multi_app_msg_item_list']:
                        title = j['title']
                        content_url = j['content_url'].replace('#wechat_redirect', '')
                        passage_list.append([local_time, create_time, title, content_url])
            print('该页包含 ' + str(len(passage_list)) + ' 篇文章')
            return {
                'm_flag': 1,
                'passage_list': passage_list,
                'length': len(passage_list)
            }
        elif '"home_page_list":[]' in res.text:
            print('\n出现：操作频繁，请稍后再试\n该号已被封禁，请解封后再来！！！\n')
            return {'m_flag': 0}
        else:
            print('请求结束！未获取到第 ' + str(page + 1) + ' 页文章列表')
            return {'m_flag': 0}

    def get_detail_list(self, access_token):
        """ 输入：access_token(从fiddler获取的链接)
            输出：无（获取的文章列表将保存在本地目录下）
            功能：
                ① 保存微信公众号文章的全部内容
                ②以excel文件形式存储，文件名设置为对应公众号的名称
        """
        # 获取该公众号名称，取公众号第一页文章列表，取第一篇文章链接
        first_link = self.access_origin_list(access_token, 1, False, False)
        if first_link:  # 获取到内容
            new_url = first_link[0][3].replace('amp;', '')
            res = requests.get(new_url, headers=self.headers, verify=False)
            nickname = re.search(r'var nickname.*"(.*?)".*', res.text).group(1)  # 公众号名称
        else:
            print('获取失败')
            return None

        # 遍历文章列表，获取各文章的详情内容
        print('开始获取公众号>>>>    ' + nickname)
        print('开始检测公众号的文章列表是否存在>>>>    ')
        official_path = self.root_path + self.official_names_head + nickname  # 公众号存储根路径
        # article_contents_path = official_path + '/' + '文章列表（article_list）_原始链接.xlsx'  # 文章内容文件路径
        article_list_path = official_path + '/' + '文章列表（article_list）_直连链接.xlsx'  # 文章列表文件路径
        if not os.path.exists(article_list_path):  # 如果文件不存在
            print('文件不存在，请检查目录文件>>>>  ' + article_list_path)
        else:
            frame_df = pd.read_excel(article_list_path)  # 读取表格内容，默认打开DataFrame对象包含第一个工作表中的数据
            error_links = []
            for index, row in frame_df.iterrows():
                single_article_url = row.iloc[4]  # 获取单文章链接
                try:
                    new_messages = self.get_detail_new(single_article_url)  # 获取单文章详情信息
                    # 存储获取到的文章详情信息
                    # 创建 or 打开表格，检查文件是否存在，判断不存在时创建表格文件
                    article_detail_path = official_path + '/' + '文章详情（article_detiles）.xlsx'  # 文章详情文件路径
                    if not os.path.exists(article_detail_path): pd.DataFrame().to_excel(article_detail_path, index=False)
                    frame_df = pd.read_excel(article_detail_path)  # 读取表格内容，默认打开DataFrame对象包含第一个工作表中的数据

                    # 将新数据转换为 DataFrame 并添加到现有 DataFrame 的末尾
                    columns = ['本地创建时间', '文章发布时间', '文章标题', '文章链接', '文章文本内容',
                               '阅读量', '点赞数', '转发数', '在看数',
                               '评论', '评论点赞']  # 列名
                    new_data_df = pd.DataFrame([new_messages], columns=columns)
                    df = pd.concat([frame_df, new_data_df], ignore_index=True)

                    # 将更新后的数据写入 Excel 文件
                    df.to_excel(article_detail_path, index=False)
                    local_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 本地时间
                    print(local_time + ' 已保存文章详情>>>> ' + new_messages[2])
                    print(local_time + ' 内容存储路径>>>> ' + article_detail_path)

                    delay_time = random.uniform(3, 6)  # 延迟时间
                    print('为预防被封禁,开始延时操作，延时时间：' + str(delay_time) + '秒')
                    time.sleep(delay_time)  # 模拟手动操作，随机延时delay_time秒，预防被封禁
                except:
                    error_links.append(row.iloc[:])
                    print('有问题的链接，文章标题为>>>>    ' + row.iloc[2])
                    article_error_path = official_path + '/' + '问题链接（error_links）.xlsx'  # 文章详情文件路径
                    if not os.path.exists(article_error_path): pd.DataFrame().to_excel(article_error_path, index=False)
                    columns = ['本地保存时间', '文章发布时间', '文章名称', '文章原始链接（直接访问会提示验证）']  # 列名
                    error_data_df = pd.DataFrame(error_links, columns=columns)
                    error_forme_df = pd.read_excel(article_error_path)  # 读取表格内容，默认打开DataFrame对象包含第一个工作表中的数据
                    dfs = pd.concat([error_forme_df, error_data_df], ignore_index=True)
                    # 将更新后的数据写入 Excel 文件
                    dfs.to_excel(article_error_path, index=False)
                    local_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())  # 本地时间
                    print(local_time + ' 已保存问题文章链接>>>> ' + row.iloc[2])
                    print(local_time + ' 内容存储路径>>>> ' + article_error_path)

    def get_detail_new(self, link):
        """ 输入：文章链接（无需验证，可直接访问）
            输出：单文章详情信息
        """
        '''获取部分请求参数'''
        contents = self.get_one_article(link, False, False)
        # nickname = re.search(r'var nickname.*"(.*?)".*', contents['content']).group(1)  # 公众号名称
        # article_link = re.search(r'var msg_link = .*"(.*?)".*', contents['content']).group(1)  # 文章短链接
        createTime = re.search(r"var createTime = '(.*?)'.*", contents['content']).group(1)  # 文章发布时间 detail_time
        # author = re.search(r'var author = "(.*?)".*', contents['content']).group(1)  # 文章作者
        article_title = re.search(r"var title = '(.*?)'.*", contents['content']).group(1)  # 文章标题
        # 将文字内容转换为列表形式存储
        soup = BeautifulSoup(contents['content'], 'html.parser')
        original_texts = soup.getText().split('\n')  # 将页面所有的文本内容提取，并转为列表形式
        article_texts = list(filter(lambda x: bool(x.strip()), original_texts))  # 列表形式的文章内容 texts
        r = ''
        for rand in range(0, 16):
            r += str(random.randint(0, 9))
        r = '0.' + r
        appmsg_type = "9"
        mid = str(link).split('mid=')[1].split('&')[0]
        sn = str(link).split('sn=')[1].split('&')[0]
        idx = str(link).split('idx=')[1].split('&')[0]
        ct = ''
        comment_id = re.search("var comment_id = '(.*?)'.*", contents['content'])
        if comment_id:
            comment_id = re.search("var comment_id = '(.*?)'.*", contents['content']).group(1)
        else:
            print('没有匹配到comment_id，文章标题为：' + article_title)
            comment_id = ''

        # version = contents['content'].split('_g.clientversion = "')[1].split('"')[0]
        if 'var req_id = ' in contents['content']:
            req_id = contents['content'].split('var req_id = ')[1].split(';')[0].replace("'", "").replace('"', '')
        else:
            print('没有匹配到req_id，文章标题为：' + article_title)
            req_id = ''
        # print(r, appmsg_type, mid, sn, idx, ct, comment_id, version, req_id, createTime, article_texts)

        '''获取文章详情信息'''
        detail_url = ('https://mp.weixin.qq.com/mp/getappmsgext?f=json&mock=&fasttmplajax=1&f=json' + '&uin=' + self.uin
                      + '&key=' + self.key + '&pass_ticket=' + self.pass_ticket + '&__biz=' + self.biz)
        data = {
            'r': r,
            'sn': sn,
            'mid': mid,
            'idx': idx,
            'req_id': req_id,
            'title': article_title,
            'comment_id': comment_id,
            'appmsg_type': appmsg_type,
            '__biz': self.biz,
            'pass_ticket': self.pass_ticket,
            'abtest_cookie': '', 'devicetype': 'Windows 7 x64', 'version': '63090b13', 'is_need_ticket': '0',
            'is_need_ad': '0', 'is_need_reward': '0', 'both_ad': '0', 'reward_uin_count': '0', 'send_time': '',
            'msg_daily_idx': '1', 'is_original': '0', 'is_only_read': '1', 'scene': '38', 'is_temp_url': '0',
            'item_show_type': '0', 'tmp_version': '1', 'more_read_type': '0', 'appmsg_like_type': '2',
            'related_video_sn': '', 'related_video_num': '5', 'vid': '', 'is_pay_subscribe': '0',
            'pay_subscribe_uin_count': '0', 'has_red_packet_cover': '0', 'album_id': '1296223588617486300',
            'album_video_num': '5', 'cur_album_id': 'undefined', 'is_public_related_video': 'NaN',
            'encode_info_by_base64': 'undefined', 'exptype': '', 'export_key_extinfo': '', 'business_type': '0',
        }
        res = requests.post(url=detail_url, data=data, headers=self.headers, cookies=self.cookies, verify=False)
        # print(res.text)
        read_num = jsonpath.jsonpath(json.loads(res.text), "$.." + "read_num")
        like_num = jsonpath.jsonpath(json.loads(res.text), "$.." + "old_like_num")
        share_num = jsonpath.jsonpath(json.loads(res.text), "$.." + "share_num")
        show_read = jsonpath.jsonpath(json.loads(res.text), "$.." + "show_read")

        # 获取评论以及评论点赞数
        comment_url = ('https://mp.weixin.qq.com/mp/appmsg_comment?action=getcomment&__biz=' + self.biz +
                       '&appmsgid=2247491372&idx=1&comment_id=' + comment_id + '&offset=0&limit=100&uin='
                       + self.uin + '&key=' + self.key + '&pass_ticket=' + self.pass_ticket
                       + '&wxtoken=&devicetype=Windows+10&clientversion=62060833&appmsg_token=')
        response = requests.get(comment_url, headers=self.headers, cookies=self.cookies, verify=False)
        json_content = json.loads(response.text)
        comments = jsonpath.jsonpath(json_content, '$..content')                    # 评论
        comments_star_nums = jsonpath.jsonpath(json_content, '$..like_num')         # 评论点赞数

        local_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())         # 本地时间
        if read_num == [] or read_num == '':
            return '', '', '', ''
        else:
            return (local_time, createTime, article_title, link, article_texts,  # 本地创建时间，文章发布时间，标题，链接，文本，
                    read_num[0], like_num[0], share_num[0], show_read[0],  # 阅读量，点赞数，转发数，在看数，
                    comments, comments_star_nums)  # 评论，评论点赞

