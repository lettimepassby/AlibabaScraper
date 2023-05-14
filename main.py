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
import ctypes
import threading
import sys



current_version = "v1.2.0.0"

total_alibaba = pd.DataFrame(columns=['标题', '主页', '图片', '主图', '价格', '属性'])

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
    alibaba['属性'] = None
    total_alibaba = pd.concat([total_alibaba, alibaba], ignore_index=True)
    # 图片保存
    from urllib import request
    img_progress['maximum'] = len(alibaba) - 1
    img_progress['value'] = 0
    for row_idx in range(len(alibaba)):
        title = alibaba.loc[row_idx]['标题']
        rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/ \ : * ? " < > |'
        new_title = re.sub(rstr, "_", title)  # 替换为下划线
        new_title = str(page) + '_' + str(row_idx) + '_' + new_title
        path = './result/'+new_title
        folder = os.path.exists(path)
        if not folder:
            os.makedirs(path)
        mainUrl = alibaba.loc[row_idx]['主图']
        mainUrl = 'https:' + mainUrl
        name = mainUrl.split('.')
        name = name[-1]
        name = 'main.' + name
        urllib.request.urlretrieve(mainUrl, path + '/' + name)
        img = alibaba.loc[row_idx]['图片']
        for ii in range(len(img)):
            imgUrl = 'https:' + img[ii]
            name = imgUrl.split('.')
            name = name[-1]
            name = str(ii) + '.' + name
            urllib.request.urlretrieve(imgUrl, path + '/' + name)
        img_progress['value'] += 1
        img_progress.update()
        productUrl = alibaba.loc[row_idx]['主页']
        productUrl = 'https:' + productUrl
        html = requests.get(productUrl,hader).text
        pagedata = re.compile(r"window.detailData = (?P<pgdata>.*?)window.detailData.scVersion", re.S)
        result = pagedata.finditer(html)
        for iii in result:
            resultjson = iii.group('pgdata')
        data = json.loads(resultjson)
        productBasicProperties = {item['attrName']: item['attrValue'] for item in data['globalData']['product']['productBasicProperties']}
        alibaba.loc[row_idx, '属性'] = json.dumps(productBasicProperties)
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

# 检测更新的函数
def check_for_updates():
    announcement, version, updates, updataUrl = get_announcement_and_version()
    if announcement and version and updates and updataUrl:
        if version == current_version:
            update_info.set("已是最新版本")
            announcement_text.config(state='normal')
            announcement_text.delete("1.0", tk.END)
            announcement_text.insert(tk.END, f"{announcement}\n")
            announcement_text.config(state='disabled')
        else:
            update_info.set(f"更新信息：{version}")
            announcement_text.config(state='normal')
            announcement_text.delete("1.0", tk.END)
            announcement_text.insert(tk.END, f"{announcement}\n\n更新内容：\n{updates}\n")
            announcement_text.config(state='disabled')
            if tk.messagebox.askyesno("更新提示", "检测到新版本，是否更新？"):
                download_and_replace(updataUrl, version, update_progress)  # 调用下载和替换函数
    else:
        update_info.set("更新信息：获取失败")
        announcement_text.config(state='normal')
        announcement_text.delete("1.0", tk.END)
        announcement_text.insert(tk.END, "无法获取公告内容\n")
        announcement_text.config(state='disabled')




def download_and_replace(updataUrl, version, download_progress):
    try:
        response = requests.get(updataUrl, stream=True)  # 设置 stream=True
        if response.status_code == 200:
            file_size = int(response.headers.get('content-length', 0))  # 获取文件大小
            download_progress['maximum'] = file_size
            download_progress['value'] = 0

            # 下载最新版本程序文件
            file_path = os.path.abspath(os.path.join(os.path.dirname(sys.executable), "AlibabaScraper.exe"))
            new_file_path = os.path.join(os.path.dirname(file_path), f"AlibabaScraper_{version}.exe")
            with open(new_file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=1024):  # 按块下载
                    if chunk:
                        f.write(chunk)
                        download_progress['value'] += len(chunk)
                        download_progress.update()

            # 提示用户新版本已下载
            tk.messagebox.showinfo("提示", f"新版本已下载，请查看程序目录 {new_file_path}")
            sys.exit()

        else:
            tk.messagebox.showinfo("提示", "下载最新版本程序失败")
    except Exception as e:
        tk.messagebox.showinfo("提示", f"下载最新版本程序失败{str(e)}")




def start_check_for_updates():
    download_progress_bar = show_download_progress()
    update_thread = threading.Thread(target=check_for_updates, args=(download_progress_bar,))  # 将进度条传递给函数
    update_thread.start()

def show_download_progress():
    download_progress_window = tk.Toplevel()
    download_progress_window.title("更新进度")
    download_progress_window.geometry("300x100")

    download_progress_label = ttk.Label(download_progress_window, text="下载进度：")
    download_progress_label.pack(pady=10)

    download_progress_bar = ttk.Progressbar(download_progress_window, orient=tk.HORIZONTAL, mode='determinate')
    download_progress_bar.pack(pady=10)

    return download_progress_bar

# 最新信息请求
def get_announcement_and_version():
    url = "https://www.fastmock.site/mock/a8e6ffc2a7e302ae0c7b6665ddccc0ea/al/check"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        announcement = data.get("announcement")
        version = data.get("version")
        updates = data.get("Updates")
        updataUrl = data.get("updataUrl")
        return announcement, version, updates, updataUrl
    else:
        return None, None, None



app = tk.Tk()
app.title("Alibaba Scraper by TimePassBy")
app.geometry("500x450")

# 创建左侧 Frame
left_frame = ttk.Frame(app, padding="10")
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# 创建右侧 Frame
right_frame = ttk.Frame(app, padding="10")
right_frame.pack(side=tk.RIGHT, fill=tk.Y)

word_label = ttk.Label(left_frame, text="请输入需要爬取的关键词：")
word_label.grid(column=0, row=0, sticky=tk.W)
word_entry = ttk.Entry(left_frame)
word_entry.grid(column=1, row=0, sticky=(tk.W, tk.E))

pages_label = ttk.Label(left_frame, text="请输入需要爬取的总页数：")
pages_label.grid(column=0, row=1, sticky=tk.W)
pages_entry = ttk.Entry(left_frame)
pages_entry.grid(column=1, row=1, sticky=(tk.W, tk.E))

start_button = ttk.Button(left_frame, text="开始爬取", command=start_crawl)
start_button.grid(column=1, row=2, sticky=tk.E)

progress_label = ttk.Label(left_frame, text="爬取总进度：")
progress_label.grid(column=0, row=3, sticky=tk.W)
progress = ttk.Progressbar(left_frame, orient=tk.HORIZONTAL, mode='determinate')
progress.grid(column=0, row=4, columnspan=2, sticky=(tk.W, tk.E))

img_progress_label = ttk.Label(left_frame, text="子进度：")
img_progress_label.grid(column=0, row=5, sticky=tk.W)
img_progress = ttk.Progressbar(left_frame, orient=tk.HORIZONTAL, mode='determinate')
img_progress.grid(column=0, row=6, columnspan=2, sticky=(tk.W, tk.E))


# 在右侧 Frame 中添加公告框
announcement_label = ttk.Label(right_frame, text="公告：")
announcement_label.pack(pady=10)

announcement_text = tk.Text(right_frame, wrap=tk.WORD, height=10, width=30, state='disabled')
announcement_text.pack(pady=10)

# 在左侧 Frame 中添加当前版本号 Label
version_label = ttk.Label(right_frame, text="当前版本号：" + current_version)
version_label.pack(pady=10)


# 在右侧 Frame 中添加更新信息 Label
update_info = tk.StringVar()
update_label = ttk.Label(right_frame, textvariable=update_info)
update_label.pack(pady=10)

# 在右侧 Frame 中添加检测更新的按钮
check_updates_button = ttk.Button(right_frame, text="检测更新", command=start_check_for_updates)
check_updates_button.pack(pady=10)

update_progress_label = ttk.Label(right_frame, text="更新进度：")
update_progress_label.pack(pady=10)
update_progress = ttk.Progressbar(right_frame, orient=tk.HORIZONTAL, mode='determinate')
update_progress.pack(pady=10)

check_for_updates()  # 在这里添加调用 check_for_updates 函数的代码

app.mainloop()
