#-*- coding:UTF-8-*-

import sys
import time
import requests
import urllib
import numpy as np
from bs4 import BeautifulSoup
import xlwt

reload(sys)
sys.setdefaultencoding('utf-8')

# UA
hds = [{'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11'},\
       {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:34.0) Gecko/20100101 Firefox/34.0'},\
       {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36 OPR/26.0.1656.60'},
       ]

def book_spider(book_tag):
    page_num = 0
    book_list = []
    try_times = 0

    while(1):
        url = 'http://book.douban.com/tag/' + urllib.quote(book_tag) + '?start=' + str(page_num*20) + '&type=T'
        time.sleep(np.random.rand()*5)

        #
        try:
            req = requests.get(url, headers=hds[page_num%len(hds)])
            html_text = req.text
        except:
            pass

        soup = BeautifulSoup(html_text, 'lxml')
        list_soup = soup.find('ul', {'class': 'subject-list'})

        try_times += 1
        if list_soup == None and try_times < 200:
            continue
        elif  len(list_soup) <= 1:
            break

        for book_info in list_soup.findAll('li'):
            info = book_info.find('h2')
            title_info = info.find('a')
            title = title_info.attrs['title'].strip()
            book_url = title_info.attrs['href'].strip()
            try:
                pub_info = book_info.find('div', {'class': 'pub'})
                pub = pub_info.string.strip()
            except:
                pub = u'暂无'
            try:
                stars_info = book_info.find('div', {'class': 'star clearfix'})
                rating_num = stars_info.find('span', {'class': 'rating_nums'}).string.strip()
                rating_per = stars_info.find('span', {'class': 'pl'}).string.strip()
            except:
                rating_num = 0
                rating_per = u'暂无'
            try:
                introduction = book_info.find('p').string.strip()
            except:
                introduction = u'暂无'
            book_list.append([title, rating_num, rating_per, pub, introduction, book_url])
            try_times = 0
        page_num += 1
        print 'Downloading Information From Page %d' % page_num
    return book_list

def spider_start(book_tag_list):
    book_lists = []
    for book_tag in book_tag_list:
        print book_tag
        book_list = book_spider(book_tag)
        book_list = sorted(book_list, key=lambda x: x[1], reverse=True)
        book_lists.append(book_list)
    return book_lists

def save_excel(book_lists, book_tag_lists):
    workbook = xlwt.Workbook()
    book_num = 0
    for book_sheet in book_tag_lists:
        worksheet = workbook.add_sheet(book_sheet.decode('utf-8'))
        sheet_title = [u'书名', u'评分', u'评分人数', u'出版信息', u'书籍简介', u'链接地址']
        i = 0
        for x in sheet_title:
            worksheet.write(0,i, x)
            i += 1
        row = 1
        for x_inbook in book_lists[book_num]:
            for col in range(6):
                worksheet.write(row, col, x_inbook[col])
            row += 1
        book_num += 1
    save_path = ''

    for b_name in book_tag_lists:
        b_list = b_name + '-'
    save_name = save_path + 'book_list' + b_list.decode('utf-8') + '.xlsx'
    workbook.save(save_name)


if __name__ == '__main__':
    book_tag_lists = ['计算机','机器学习','linux','android','数据库','互联网']
    book_lists = spider_start(book_tag_lists)
    save_excel(book_lists, book_tag_lists)