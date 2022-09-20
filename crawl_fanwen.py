#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/25'

from gevent import monkey
monkey.patch_all()


import os
import re
import time

import requests
from selenium_helper import SeleniumHelper
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from pyquery import PyQuery as pq
from selenium_helper import SeleniumHelper
import gevent



import config as cfg
import utils
import common as com


test_url='https://www.diyifanwen.com/fanwen/maimaihetong/3384933.html'

base_url='https://www.diyifanwen.com'

menu_url='{}/fanwen/maimaihetong/index_75.html'.format(base_url)

detail_url='https://s.diyifanwen.com/down/?action=print&id={}/fanwen/hetongyangben/2745634.htm'.format(base_url)




def get_detail_li(url):
    li=[]
    res=requests.get(url,headers=cfg.get_headers()).content.decode('gbk')
    html=pq(res)
    datas=html('#AListBox>ul>li>a')
    for data in datas.items():
        title=data.text()
        url=data.attr('href')
        if not url.startswith('http:'):
            url='https:'+url
        url='https://s.diyifanwen.com/down/?action=print&id={}'.format(url)
        print(title)
        print(url)
        li.append((title,url))
    return li



def get_print_text(browser,url):
    browser.get(url)
    title_css='div#title'
    wait=WebDriverWait(browser,10)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,title_css)))
    time.sleep(2)
    li=[]

    html=pq(browser.page_source)
    title=html('#title').text()
    title='<h1>{}</h1>'.format(title)
    datas=html('div#title+ div>p')
    li.append(title)
    for data in datas.items():
        content=data.html()
        if content:
            content=content.replace('　　','\n\n')
            li.append(content)
    return li


# def get_detail_url(url):
#     li=[]
#     res=requests.get(url,headers=cfg.get_headers()).content
#     html=pq(res)
#
#     # print(html.text())
#     title=html('#ArtContent>h1').text()
#     datas=html('#ArtContent>p:not(last-of-type)')
#     title='<h1>{}</h1>'.format(title)
#     if '_' in url:
#         title=''
#
#     print('-'*30)
#     li.append(title)
#
#     for num,data in enumerate(datas.items()):
#         text=data.html()
#         # if num>=2:
#         words=re.findall('<a.+?</a>',text)
#         word_2=re.findall('来.+?\.htm|本.+?\.htm|文.+?\.htm',text,re.S)
#         word_3=['声明：仅供参考!']
#         words.extend(word_2)
#         words.extend(word_3)
#         if words!=[]:
#             for word in words:
#                 if word in text:
#                     text=text.replace(word,'')
#         text=text.replace('　　','\n\n')
#         if '本文地址' in text:
#             break
#         # print(text)
#         li.append(text)
#
#
#     return li
#
#
# def handle_many(url):
#     res=requests.get(url,headers=cfg.get_headers()).text
#     html=pq(res)
#     is_many_page=re.findall('TxtPart',res)
#     # print(is_many_page)
#
#     if is_many_page!=[]:
#         # print('is_many_page:{}'.format(is_many_page))
#         li=[]
#         page=html('div#TxtPart>span:first-of-type>font.red:first-of-type').text()
#         if not page:
#             page=html('div.TxtPart:eq(0)>span:first-of-type>font.red:first-of-type').text()
#         # print('page:{}'.format(page))
#
#         for n in range(1,int(page)+1):
#             if n==1:
#                 page_url=url
#             else:
#                 page_url=url.replace('.htm','')+'_{}.htm'.format(n)
#
#             res=get_detail_url(page_url)
#             # print(res)
#             li.extend(res)
#         return li
#     else:
#         res=get_detail_url(url)
#         return res



def main(browser,page):


    # 合同样本78 从大到小
    # menu_url='https://www.diyifanwen.com/fanwen/hetongyangben/index_{}.html'.format(page)

    # 采购合同
    # menu_url='https://www.diyifanwen.com/fanwen/caigouhetong/index_{}.html'.format(page)

    # 租房合同
    # menu_url='https://www.diyifanwen.com/fanwen/zufanghetong/index_{}.html'.format(page)

    # 技术合同
    # menu_url='https://www.diyifanwen.com/fanwen/jishuhetong/index_{}.html'.format(page)

    # 代理合同
    menu_url='https://www.diyifanwen.com/fanwen/dailihetong/index_{}.html'.format(page)

    detail_li=get_detail_li(menu_url)

    for title,detail_url in detail_li:
        try:
            downloaded=utils.get_changed(com.path,need_ext=False)
            if title not in downloaded:
                text_li=get_print_text(browser,detail_url)
                file_name=com.path+'/'+title+'.doc'
                com.save_file(com.path,file_name,text_li)
        except Exception as e:
            print(e)


if __name__=='__main__':
    # browser=SeleniumHelper().browser
    # for i in range(42,46):
    #     main(browser,i)
    # browser.close()


    # utils.doc_to_docx(com.path,del_origin_file=True)
    # com.modify_docx(com.path)

    # utils.rename_file(cfg.save_path,cfg.word_filter)
    # utils.rename_file_front(com.path)
    # utils.rename_file_end(com.path)

    utils.many_doc_to_pdf(cfg.save_path)
    utils.many_pdf_to_png(cfg.save_path,cfg.pic_path)
