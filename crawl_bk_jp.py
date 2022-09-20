#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/16'

import os,re
import time
import threading

import requests
from pyquery import PyQuery as pq
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


import config as cfg
from selenium_helper import SeleniumHelper
from baiduwenku import write_to_file

base_url='https://wenku.baidu.com'

# menu_url='{}/ndbgc/org/legal?pn=100&sort=1'.format(base_url)

# detail_url='https://wenku.baidu.com/view/51dbb2f6de80d4d8d05a4f7a'

# detail_url='https://wenku.baidu.com/view/a6d88a73a8956bec0975e356.html'

detail_url='https://wenku.baidu.com/view/6eb30d73915f804d2b16c1c8'

li=[]
li_2=[]


def get_detail_url(browser,url):

    browser.get(url)
    wait=WebDriverWait(browser,10)
    title_css='.gd-bd-title'
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,title_css)))
    # title_ele=browser.find_element_by_css_selector(title_css)
    # print(browser.page_source)
    return browser.page_source

def get_detial_url_2(url):
    res=requests.get(url,headers=cfg.get_headers()).content.decode('gbk')
    return res


def get_data_li(page_source):
    html=pq(page_source)
    datas=html('p.article-title')
    for i in datas.items():
        if i('span.ic-doc'):
            url=i('a').attr('href')
            url=base_url+url
            title=i('a').attr('title')
            print(title)
            # print(url)
            li.append((title,url))


def download_file(browser,url,title):
    browser.get(url)
    time.sleep(0.5)
    res=browser.page_source
    is_share=is_share_doc(res)
    if is_share:
        print('{}：是共享文档'.format(title))
        li_2.append(url)
        # file_name=title+'.doc'
        # write_to_file(browser,url,"word",file_name)
    else:
        print('{}：是其他文档'.format(title))


def download_file_direct(browser,url,title):
    browser.get(url)
    time.sleep(0.5)
    res=browser.page_source
    is_share=is_share_doc(res)
    if is_share:

        print('{}：是共享文档'.format(title))
        res=browser.page_source
        html=pq(res)
        content=html('#pageNo-1')
        print(content.html())


    else:
        print('{}：是其他文档'.format(title))


def is_share_doc(page_source):
    html=pq(page_source)
    doc_type=html('div.doc-tag-wrap>div[style="display: block;"]>span').text()

    return doc_type=='共享文档'







def save_file(path,li):
    with open(path,'a') as f:
        for i in li:
            f.write(i)
            f.write('\n')
    print('写入完毕')

# def handle_doc_type(url):
#     if is_share_doc(url):
#             li_2.append(url)

def main():
    # menu_url='https://wenku.baidu.com/jingpin/195?ca=195&od=2&pn=6'




    for i in range(0,40):
        menu_url='{}/jingpin/75?ca=75&od=2&pn={}'.format(base_url,i*6)
        print(menu_url)

        page_source=get_detial_url_2(menu_url)
        get_data_li(page_source)

    browser=SeleniumHelper().browser


    for title,url in li:
        download_file(browser,url,title)

    save_file('2.txt',li_2)
    # time.sleep(20)

if __name__=='__main__':
    # main()

    url='https://wenku.baidu.com/view/7f637d7e6d175f0e7cd184254b35eefdc9d3150f'
    browser=SeleniumHelper().browser
    download_file_direct(browser,url,'测试1')
