# -*- coding: utf-8 -*-
"""
Created on Thu Apr 12 16:04:56 2018

@author: LYC
"""

############# 抓PTT 新竹版 ######################################
import requests
from bs4 import BeautifulSoup
import time
import threading
import sys
import xlwt
import xlrd

book = xlwt.Workbook(encoding = "utf-8")
sheet1 = book.add_sheet("贈送")


def main(orig_args):
    filename = "PTT新竹版贈送.xls"
    output(filename)

def output(filename):
    sheet1.write(0,0,'編號')
    sheet1.write(0,1,'日期')
    sheet1.write(0,2,'物品')
    sheet1.write(0,3,'網址')
    book.save(filename)

main(sys.argv)   

def open_save(pos1,pos2,str):
    try:
        inbook = xlrd.open_workbook('PTT新竹版贈送.xls',formatting_info = True)
        outbook = copy(inbook)
        outbook.get_sheet(0).write(pos1,pos2,str)
        outbook.save('PTT新竹版贈送.xls')
    except IOError:
        print('ERROR!')
        sys.exit('No such file: PTT新竹版贈送.xls')


def get_page(url):
    res = requests.get(url)
    return res.text

def get_next_url(dom):
    soup = BeautifulSoup(dom, 'lxml')
    base_url = 'https://www.ptt.cc'
    next_url =  []
    check = []
    divs = soup.find_all('div', 'btn-group btn-group-paging')
    for d in divs:
        #if d.find('a').string == ' 上頁':
            #re.search(r'href="/bbs/Beauty/index2191.html">&lsaquo; 上頁)
#            print(d)
#            print()
#            print(d.find_all('a', 'btn wide')[1])
            check = d.find_all('a', 'btn wide disabled')
            href = d.find_all('a', 'btn wide')[1]['href']
            next_url.append(href)
    if check and check[0].text == '‹ 上頁':
        print(check[0].text)
        return None
#    print(base_url + next_url[0])
    
    return base_url + next_url[0]
    

def get_gift(dom,number):
    #print(res.text)
    soup = BeautifulSoup(dom, "lxml")
    
    list_title = []
    list_date = []
    list_author = []
    list_link = []
    base_url = 'https://www.ptt.cc'
    
    articles = []
    
    divs = soup.find_all('div', 'r-ent')
    
    

    date = time.strftime("%m/%d")
    for d in divs:
#        if '0' + d.find('div', 'date').string.lstrip() == date:  ##### 只抓今天的 ######
    #        push_count = 0
    #        if d.find('div', 'nrec').string:
    #            try:
    #                push_count = int(d.find('div','nrec').string)
    #            except ValueError:
    #                pass
            date_article = d.find('div', 'date').string.lstrip()
            if d.find('a'):
                href = d.find('a')['href']
                title = d.find('a').string
                articles.append({
    #                'push_count': push_count,
                    'title': title,
                    'href': base_url + href,
                    'date': date_article
                    })
           
    ##############  DEMO
#    print(articles)
#    print(len(articles))
    
    gift = '贈送'
    for ord in range(len(articles)):
        # 確認為贈送
        if articles[ord]['title'][1:3] == gift: # [贈送]
            # 避免跟list中的重複
            Judge = 0 
            ord_list = 0
            for ord_list in range(len(list_title)):
                #重複  
#                print(articles[ord]['title'][0:len(articles[ord]['title'])])
#                print(list_title)
                if len(list_title)!=0 and articles[ord]['title'][0:len(articles[ord]['title'])] == list_title[ord_list]:
                    Judge = -1
            #沒重複        
            if Judge == 0:
                list_title.append(articles[ord]['title'][0:len(articles[ord]['title'])])
                list_link.append(articles[ord]['href'][0:len(articles[ord]['href'])])
                list_date.append(articles[ord]['date'][0:len(articles[ord]['date'])])
            else:
                Judge = 0
                
    # 存入Excel
    for ord_list in  range(len(list_title)):
            
            open_save(ord_list+1 + number ,0,ord_list+1 + number)
            open_save(ord_list+1 + number ,1,list_date[ord_list])
            open_save(ord_list+1 + number ,2,list_title[ord_list])
            open_save(ord_list+1 + number ,3,list_link[ord_list])
    
    next_order = number + len(list_title)
    return next_order
    
#    print(list_title,list_link,list_date)

############  固定十秒跑一次  #############
#def t2():
#    while 1:
#        get_gift()
#        time.sleep(10)
#
#t = threading.Thread(target = t2)
#t.start()

#get_gift()


def percentage(times):
    print('完成度: ''%.2f%%' % (times/4004*100))


urls = 'https://www.ptt.cc/bbs/Hsinchu/index.html'
times = 1;
next_order = 0
while urls != None:
    res = get_page(urls)
    next_order = get_gift(res,next_order)
#    print('next_order = ', next_order)
    urls = get_next_url(res)
    percentage(times)
    times = times + 1

print('Done!!!!!!!!!!!!!!!!!!')