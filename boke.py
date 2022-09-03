# -*- coding = UTF-8 -*-
# @Time : 2021/6/21 15:12
# @Author : 伏珊瑞
# @File : boke.py
# @Software : PyCharm
import requests
import xlwt
from bs4 import BeautifulSoup
import urllib.request

url_list1 = []  # 用于存放标题和url
url_list2 = []  # 用于存放日期
url_list3 = []  # 用于存放详情界面


# 获取源码
def get_content(url):
    html = requests.get(url).content
    return html


# 获取列表页中的所有博客url，标题，链接，日期；
def get_url(html):
    soup = BeautifulSoup(html, 'html.parser')  # lxml是解析方式，第三方库

    blog_url_list1 = soup.find_all('div', class_='postTitle')
    for i in blog_url_list1:
        url_list1.append([i.find('a').text, i.find('a')['href']])

    # 获取列表日期，时间
    blog_url_list2 = soup.find_all('div', class_='postDesc')
    for i in blog_url_list2:
        s = i.text[9:19]
        # print(s)
        url_list2.append(s)


# 开始爬取
for page in range(1, 7):  # 定义要爬取的页面数
    url = 'http://www.cnblogs.com/wrljzb/default.html?page={}'.format(page)
    # print(url)
    get_url(get_content(url))

# 详情页中内容进行分步爬取
lens = len(url_list1)

for j in range(0, lens):
    url = url_list1[j][1]
    req = urllib.request.Request(url)
    resp = urllib.request.urlopen(req)
    html_page = resp.read().decode('utf-8')
    soup = BeautifulSoup(html_page, 'html.parser')

    # print(soup.prettify())

    div = soup.find(id="cnblogs_post_body")

    url_list3.append([div.get_text()])

newTable = 'ji.xls'  # 生成的excel名称
wb = xlwt.Workbook(encoding='utf-8')  # 打开一个对象
ws = wb.add_sheet('blog')  # 添加一个sheet
headData = ['博客标题', '链接', '时间', '详细内容']
# 写标题
for colnum in range(0, 4):  # 定义的四列名称
    ws.write(0, colnum, headData[colnum], xlwt.easyxf('font:bold on'))  # 第0行的第colnum列写入数据headDtata[colnum]，就是表头，加粗
index = 1
lens = len(url_list1)
# 写内容
# print(len(url_list1), len(url_list2))
print(url_list2)
for j in range(0, lens):
    ws.write(index, 0, url_list1[j][0])
    ws.write(index, 1, url_list1[j][1])
    ws.write(index, 2, url_list2[j])
    ws.write(index, 3, url_list3[j][0])
    index += 1  # 下一行
wb.save(newTable)  # 保存
