#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/10'


import time
import os
import shutil
import json

from selenium_helper import SeleniumHelper
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import requests

from config import save_path,fm_headers,JOB_TYPE,USEFUL_TYPE,COMMON_TYPE,PERSON_TYPE,pic_path,YEAR,MONTH
from utils import win_handle,get_file_path


def word_cls_handle(file_name):
    for a in JOB_TYPE:
        if a in file_name:
            return "校园相关"

    for b in COMMON_TYPE:
        if b in file_name:
            return "常用范文"

    for c in USEFUL_TYPE:
        if c in file_name:
            return "实用文档"

    for e in PERSON_TYPE:
        if e in file_name:
            return '人事管理'

    return '实用文档'


upload_url='https://www.feimaoyun.com/member.html#/itemword'


def get_uploaded(start,end):
    li=[]
    for i in range(start,end):
        url='https://www.feimaoyun.com/index.php/choice_getoffice'

        data=dict(pg=i,status='')
        res=requests.post(url,headers=fm_headers,data=data).text

        res=json.loads(res)
        datas=res['data']['offices']

        for data in datas:
            ppt_name=data['name']
            if ppt_name.endswith(('.docx','.doc')):
                li.append(ppt_name)
    li=list(set(li))
    return li

def need_upload_file(start,end):
    has_uploaded_li=get_uploaded(start,end)
    for root,dirs,files in os.walk(save_path):
        for f in files:
            if f in has_uploaded_li:
                full_path=root+'/'+f
                dst_path='g:/资源/WORD资源/{}/{}/{}'.format(YEAR,MONTH,f)
                shutil.move(full_path,dst_path)
    print('过滤完成！')


def send_to_feemoo(browser,path):
    browser.get(upload_url)
    upload_css='div.upload-top-nav>span.upload-btn'

    time.sleep(1)

    upload_ele=browser.find_element_by_css_selector(upload_css)
    upload_ele.click()

    word_css='div.content>div.sub-cld:nth-of-type(1)'

    wait=WebDriverWait(browser,10)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,word_css)))
    ppt_ele=browser.find_element_by_css_selector(word_css)
    ppt_ele.click()


    upload_file_css='.dt-file'
    upload_ele=browser.find_element_by_css_selector(upload_file_css)
    upload_ele.click()
    time.sleep(0.5)
    win_handle(path)
    time.sleep(0.5)

    next_step_css='div.step-box1 span.next-btn'
    next_step_ele=browser.find_element_by_css_selector(next_step_css)
    next_step_ele.click()
    time.sleep(1)

    _,file_name=os.path.split(path)
    word_type=word_cls_handle(file_name)

    choice_css='div.step-box2 div.el-select'
    choice_ele=browser.find_element_by_css_selector(choice_css)
    choice_ele.click()
    time.sleep(0.5)

    select_target_xpath='//li/span[contains(text(),"{}")]'.format(word_type)
    select_file_ele=browser.find_element_by_xpath(select_target_xpath)
    select_file_ele.click()

    png_dir_name=file_name.replace('.docx','')
    png_dir_name=png_dir_name.replace('.doc','')

    png_paths='{}/{}'.format(pic_path,png_dir_name)
    png_paths=get_file_path(png_paths,end_with=('.png'),to_win32=True)

    for png_path in png_paths:
        time.sleep(0.5)
        upload_png_css='#softimgupload>span.el-icon-plus'
        upload_png_ele=browser.find_element_by_css_selector(upload_png_css)
        upload_png_ele.click()
        time.sleep(0.5)
        win_handle(png_path)
        # if len(png_paths)<=2:
        #     time.sleep(1)
        # else:
        time.sleep(0.5)

    if len(png_paths) < 2:
        time.sleep(2)

    submit_css='div.step-box2 span.next-btn'
    submit_ele=browser.find_element_by_css_selector(submit_css)
    submit_ele.click()

    # 小于5M等待6秒
    if os.path.getsize(path)>=200000000:
        time.sleep(60)
    if os.path.getsize(path)>=100000000:
        time.sleep(40)
    if os.path.getsize(path)>=40000000:
        time.sleep(20)
    if os.path.getsize(path)>=20000000:
        time.sleep(15)
    elif os.path.getsize(path)>=5000000:
        time.sleep(10)
    else:
        time.sleep(6)

    dst_path='g:/资源/WORD资源/{}/{}/{}'.format(YEAR,MONTH,file_name)

    shutil.move(path,dst_path)

    # print('{}:上传成功'.format(file_name))

    browser.get('https://www.baidu.com/')
    time.sleep(2)



def send_to_new_feemoo(browser,path):
    browser.get(upload_url)
    upload_css='#softdocupload'

    time.sleep(1)

    upload_ele=browser.find_element_by_css_selector(upload_css)
    upload_ele.click()


    time.sleep(1 )
    win_handle(path)
    time.sleep(0.5)



    _,file_name=os.path.split(path)
    word_type=word_cls_handle(file_name)


    select_target_xpath='//a[contains(text(),"{}")]'.format(word_type)
    select_file_ele=browser.find_element_by_xpath(select_target_xpath)
    select_file_ele.click()

    png_dir_name=file_name.replace('.docx','')
    png_dir_name=png_dir_name.replace('.doc','')

    png_paths='{}/{}'.format(pic_path,png_dir_name)
    png_paths=get_file_path(png_paths,end_with=('.png'),to_win32=True)

    for png_path in png_paths:
        time.sleep(0.5)
        upload_png_css='#softimgupload>span.el-icon-plus'
        upload_png_ele=browser.find_element_by_css_selector(upload_png_css)
        upload_png_ele.click()
        time.sleep(0.5)
        win_handle(png_path)
        # if len(png_paths)<=2:
        #     time.sleep(1)
        # else:
        time.sleep(0.5)

    if len(png_paths) < 2:
        time.sleep(1)

    submit_css='div.btn-box span.next-btn'
    submit_ele=browser.find_element_by_css_selector(submit_css)
    submit_ele.click()

    # 小于5M等待6秒
    if os.path.getsize(path)>=200000000:
        time.sleep(60)
    if os.path.getsize(path)>=100000000:
        time.sleep(40)
    if os.path.getsize(path)>=40000000:
        time.sleep(20)
    if os.path.getsize(path)>=20000000:
        time.sleep(15)
    elif os.path.getsize(path)>=5000000:
        time.sleep(10)
    else:
        time.sleep(4)

    dst_path='g:/资源/WORD资源/{}/{}/{}'.format(YEAR,MONTH,file_name)

    shutil.move(path,dst_path)

    # print('{}:上传成功'.format(file_name))

    browser.get('https://www.baidu.com/')
    time.sleep(1)
    # time.sleep(2)


def main():
    s1=SeleniumHelper()
    browser=s1.browser
    list_path=get_file_path(save_path,to_win32=True)
    count=len(list_path)
    for i in list_path:
        # send_to_feemoo(browser,i)
        send_to_new_feemoo(browser,i)
        count-=1
        print('剩余{}：{}'.format('word',count))

if __name__=='__main__':
    # 检查遗漏的
    # need_upload_file(1,11)


    main()
