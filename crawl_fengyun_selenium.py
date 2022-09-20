#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/11/6'


import os,shutil,time,re,random,zipfile


import requests
from pyquery import PyQuery as pq
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys


from selenium_helper import SeleniumHelper,SeleniumHelperNew
import config as cfg
import common as com
import utils
import crawl_fengyun as cf

base_url='http://www.ppt118.com'

menu_url='http://www.ppt118.com/words/list/t122/r3_5.html'

detail_url='http://www.ppt118.com/words/83c25ff.html'


def get_detail_li(url):
    res=requests.get(url,headers=cfg.get_headers()).content.decode('utf-8')
    html=pq(res)
    datas=html('li.listItem')

    li=[]

    for data in datas.items():
        title=data('h4.imgtt>a').text().replace(' ','')
        detail_url=data('h4.imgtt>a').attr('href')
        detail_url=base_url+detail_url

        download_uuid=detail_url.replace('.html','').split('/')[-1]

        li.append((title,detail_url,download_uuid))
    return li


def rename_fengyun(path,start,end):
    li_total=[]
    for i in range(start,end):
        # 合同协议
        # menu_url='http://www.ppt118.com/words/list/t122/r3_{}.html'.format(i)

        # 财务管理
        menu_url='http://www.ppt118.com/words/list/t121/r1_{}.html'.format(i)
        li=get_detail_li(menu_url)
        li_total.extend(li)
    print(li_total)

    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith('.zip'):
                file_name=f.replace('.zip','')
                full_path=root+'/'+f
                for title,_,uuid in li_total:
                    if file_name==uuid:
                        new_path=root+'/'+title+'.zip'
                        if os.path.exists(new_path):
                            r=str(random.randint(1,1000))
                            new_path=root+'/'+title+r+'.zip'
                        os.rename(full_path,new_path)
                        print('{}:改名成功!'.format(title+'.zip'))

def get_detail_page(browser,url):
    browser.get(url)

    download_first_css='div.r1>a.d_down'
    wait=WebDriverWait(browser,10)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,download_first_css)))
    download_first_btn=browser.find_element_by_css_selector(download_first_css)
    download_first_btn.click()

    download_now_css='a.anniu.jb.btn_download_now'
    # wait=WebDriverWait(browser,10)
    # wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,download_now_css)))
    time.sleep(1)
    js1="""
        var d=document.
        getElementsByClassName('{}')[0].
        setAttribute('target','_self')
        """.format('btn_download_now')
    browser.execute_script(js1)
    time.sleep(0.5)

    download_now_ele=browser.find_element_by_css_selector(download_now_css)
    download_now_ele.click()

    ActionChains(browser).key_down(Keys.CONTROL).send_keys("w").key_up(Keys.CONTROL).perform()
    time.sleep(5)



def main(start,end):
    for i in range(start,end):
        # 合同协议
        menu_url='http://www.ppt118.com/words/list/t122/r3_{}.html'.format(i)

        # 财务管理
        menu_url='http://www.ppt118.com/words/list/t121/r1_{}.html'.format(i)

        detail_li=get_detail_li(menu_url)
        s1=SeleniumHelperNew()
        browser=s1.base_browser()
        for title,detail_url,uuid in detail_li:
            d_path='c:/Users/rhj231223/Downloads'
            downloaded=utils.get_changed(d_path,need_ext=False)
            if uuid not in downloaded:
                get_detail_page(browser,detail_url)

    rename_fengyun(com.path,start,end)

if __name__=='__main__':

    # 合同
    # main(13,14)

    # 财务管理
    # main(1,2)

    # rename_fengyun(com.path,1,2)
    cf.compress_zip(com.path,remove_source=True)
    utils.doc_to_docx(com.path,del_origin_file=True)
