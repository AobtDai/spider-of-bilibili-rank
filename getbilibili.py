import requests
from bs4 import BeautifulSoup
import xlwt
import os

book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('text', cell_overwrite_ok=True)
sheet.write(0,0,'序号')
sheet.write(0,1,'视频名字')
sheet.write(0,2,'视频地址')
sheet.write(0,3,'Up主')
sheet.write(0,4,'Up主页地址')
sheet.write(0,5,'播放量')
sheet.write(0,6,'评论数')
sheet.write(0,7,'综合得分')
n = 1

def get_upname(view):
    upname = str(view.find(class_ = "data-box up-name"))
    re1 = '<span class="data-box up-name"><i class="b-icon author"></i>'
    re2 = "</span>"
    upname = upname.replace(str(re1), '')
    upname = upname.replace(str(re2), '')
    return upname.strip()

def get_upurl(view):
    upurl = str(view.find(class_ = "detail").find('a').get('href'))
    return upurl

def get_viewname(view):
    viewname = str(view.find(class_ = "info").find('a').string)
    return viewname

def get_viewurl(view):
    viewurl = str(view.find(class_ = "info").find('a').get('href'))
    return viewurl

def get_viewamount(view):
    viewa = ''
    data = view.find_all(class_ = "data-box")
    if len(data):
        viewa = str(data[0])
    re1 = '<span class="data-box"><i class="b-icon play"></i>'
    re2 = "</span>"
    viewa = viewa.replace(str(re1),'')
    viewa = viewa.replace(str(re2),'')
    return viewa.strip()

def get_commentamount(view):
    viewa = ''
    data = view.find_all(class_ = "data-box")
    if len(data):
        viewa = str(data[1])
    re1 = '<span class="data-box"><i class="b-icon view"></i>'
    re2 = "</span>"
    viewa = viewa.replace(str(re1),'')
    viewa = viewa.replace(str(re2),'')
    return viewa.strip()

def get_score(view):
    score = str(view.find(class_ = "pts").find('div').string)
    return score

def spiderBili(soup):
    viewlist = soup.find_all('li')
    for view in viewlist:
        if str(view.find(class_ = "data-box up-name")) == 'None':
            continue
        upname = get_upname(view)
        upurl = get_upurl(view)
        viewname = get_viewname(view)
        viewurl = get_viewurl(view)
        viewa = get_viewamount(view)
        commenta = get_commentamount(view)
        score = get_score(view)

        global n
        print("Crawling Num"+str(n)+" view\n")
        sheet.write(n,0,str(n))
        sheet.write(n,1,viewname)
        sheet.write(n,2,viewurl)
        sheet.write(n,3,upname)
        sheet.write(n,4,upurl)
        sheet.write(n,5,viewa)
        sheet.write(n,6,commenta)
        sheet.write(n,7,score)

        n = n+1

if __name__ == '__main__':
    url = "https://www.bilibili.com/v/popular/rank/all"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.164 Safari/537.36'}
    response = requests.get(url, headers = headers)
    if response.status_code==200:
        html = response.text
        soup = BeautifulSoup(html, "lxml")
        spiderBili(soup)

sheet.col(0).width = 1000
sheet.col(1).width = 11000
sheet.col(3).width = 5500
if os.path.exists("./bilibilirank.xlsx"):
  os.remove("./bilibilirank.xlsx")
book.save('bilibilirank.xlsx')
print("Excel has been created.\nPress Enter to Continue...")
input()

