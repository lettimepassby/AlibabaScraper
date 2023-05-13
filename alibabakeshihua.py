import requests
import re
import json
import pandas as pd
import openpyxl
import os
import time
import urllib.request
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

total_alibaba = pd.DataFrame(columns=['标题', '主页', '图片', '主图', '价格'])

def alibaba(word, page):
    global total_alibaba
    # url = "https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&CatId=&tab=all&SearchText=Hydraulic+Manifold+Block"
    url = "https://www.alibaba.com/trade/search?spm=a2700.galleryofferlist.0.0.69766994Z90MSS&fsb=y&IndexArea=product_en&keywords=" + word + "&tab=all&viewtype=L&&page=" + str(page)
    hader = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.52'
    }
    html = requests.get(url,hader).text
    pagedata = re.compile(r"window.__page__data__config = (?P<pgdata>.*?)window.__page__data = window.__page__data__config.props", re.S)
    result = pagedata.finditer(html)
    for i in result:
        resultjson = i.group('pgdata')
    data = json.loads(resultjson)
    # 标题提取
    title = []
    # 主页提取
    productUrl = []
    # 图片提取
    image = []
    # 主图提取
    mainImage = []
    # 价格提取
    price = []
    for i in range(len(data['props']['offerResultData']['offerList'])):
        title.append(data['props']['offerResultData']['offerList'][i]['information']['puretitle'])
        productUrl.append(data['props']['offerResultData']['offerList'][i]['information']['productUrl'])
        image.append(data['props']['offerResultData']['offerList'][i]['image']['multiImage'])
        mainImage.append(data['props']['offerResultData']['offerList'][i]['image']['mainImage'])
        price.append(data['props']['offerResultData']['offerList'][i]['promotionInfoVO']['localOriginalPriceRangeStr'])
    alibaba = pd.DataFrame({'标题':title, '主页':productUrl, '图片':image, '主图':mainImage, '价格':price})
    total_alibaba = pd.concat([total_alibaba, alibaba], ignore_index=True)
    # 图片保存
    from urllib import request
    img_progress['maximum'] = len(alibaba) - 1
    img_progress['value'] = 0
    for i in range(len(alibaba)):
        title = alibaba.loc[i]['标题']
        rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/ \ : * ? " < > |'
        new_title = re.sub(rstr, "_", title)  # 替换为下划线
        new_title = str(page) + '_' + str(i) + '_' + new_title
        path = './result/'+new_title
        folder = os.path.exists(path)
        if not folder:
            os.makedirs(path)
        mainUrl = alibaba.loc[i]['主图']
        mainUrl = 'https:' + mainUrl
        name = mainUrl.split('.')
        name = name[-1]
        name = 'main.' + name
        urllib.request.urlretrieve(mainUrl, path + '/' + name)
        img = alibaba.loc[i]['图片']
        for ii in range(len(img)):
            imgUrl = 'https:' + img[ii]
            name = imgUrl.split('.')
            name = name[-1]
            name = str(ii) + '.' + name
            urllib.request.urlretrieve(imgUrl, path + '/' + name)
        img_progress['value'] += 1
        img_progress.update()
    return True


def start_crawl():
    global total_alibaba
    if not os.path.exists('result'):
        os.makedirs('result')
    word_input = word_entry.get().replace(" ", "+")
    total_pages = int(pages_entry.get()) + 1

    progress['maximum'] = total_pages - 1
    progress['value'] = 0

    for i in range(1, total_pages):
        alibaba(word_input, i)
        progress['value'] += 1
        progress.update()
        if i < total_pages - 1:  # 只在非最后一页休眠
            time.sleep(5)
    total_alibaba.to_excel('total_alibaba.xlsx', sheet_name='Sheet1', index=False)
    tk.messagebox.showinfo("提示", "爬取数据完成。")

app = tk.Tk()
app.title("Alibaba Scraper by TimePassBy")

frame = ttk.Frame(app, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

word_label = ttk.Label(frame, text="请输入需要爬取的关键词：")
word_label.grid(column=0, row=0, sticky=tk.W)
word_entry = ttk.Entry(frame)
word_entry.grid(column=1, row=0, sticky=(tk.W, tk.E))

pages_label = ttk.Label(frame, text="请输入需要爬取的总页数：")
pages_label.grid(column=0, row=1, sticky=tk.W)
pages_entry = ttk.Entry(frame)
pages_entry.grid(column=1, row=1, sticky=(tk.W, tk.E))

start_button = ttk.Button(frame, text="开始爬取", command=start_crawl)
start_button.grid(column=1, row=2, sticky=tk.E)

progress_label = ttk.Label(frame, text="爬取总进度：")
progress_label.grid(column=0, row=3, sticky=tk.W)
progress = ttk.Progressbar(frame, orient=tk.HORIZONTAL, mode='determinate')
progress.grid(column=0, row=4, columnspan=2, sticky=(tk.W, tk.E))

img_progress_label = ttk.Label(frame, text="图片下载进度：")
img_progress_label.grid(column=0, row=5, sticky=tk.W)
img_progress = ttk.Progressbar(frame, orient=tk.HORIZONTAL, mode='determinate')
img_progress.grid(column=0, row=6, columnspan=2, sticky=(tk.W, tk.E))

app.mainloop()
