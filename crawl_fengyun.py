#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/11/2'

import os,shutil,time,re,random,zipfile

import requests
from pyquery import PyQuery as pq

import config as cfg
import common as com
import utils

path='g:/资源/WORD资源/2019/11月/测试'

base_url='http://www.ppt118.com/'

menu_url='http://www.ppt118.com/words/list/t122/r1_1.html'

detail_url='http://www.ppt118.com/words/0eaf0e43.html'

download_url='http://www.ppt118.com/download?id=0eaf0e43'


def remove_empty_file(path):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith('.zip'):
                full_path=root+'/'+f
                is_zip=zipfile.is_zipfile(full_path)
                if not is_zip:
                    os.remove(full_path)
                    print('{} 是无效文件 已删除'.format(f))

def get_detail_li(url):
    res=requests.get(url,headers=cfg.get_headers()).content.decode('utf-8')
    html=pq(res)
    datas=html('li.listItem')

    li=[]

    for data in datas.items():
        title=data('h4.imgtt>a').text().replace(' ','')
        detail_url=data('h4.imgtt>a').attr('href')
        detail_url=detail_url.replace('.html','')
        download_uuid=detail_url.replace('.htm','').split('/')[-1]
        download_url='http://www.ppt118.com/download?id={}'.format(download_uuid)

        li.append((title,download_url))
    return li

def download_file(url,file_name,use_cookies=True):
    #cookies=cfg.fengyun_cookies
    if use_cookies==True:
        res=requests.get(url,headers=cfg.common_headers,cookies=cfg.fengyun_cookies).content
    else:
        res=requests.get(url,headers=cfg.common_headers).content


    with open(file_name,'wb') as f:
        f.write(res)

    print('{}:下载成功！'.format(file_name))

def rename_file(path,dir_name,need_name):
    for root,dirs,files in os.walk(dir_name):
        for f in files:
            if f.endswith(('.docx','.doc')):
                full_path=root+'/'+f
                ext=os.path.splitext(f)[1]
                new_path=root+'/'+need_name+ext

                os.rename(full_path,new_path)
                move_after_path=path+'/'+need_name+ext
                if os.path.exists(move_after_path):
                    move_after_path=path+'/'+need_name+str(random.randint(1,1000))+ext
                shutil.move(new_path,move_after_path)
                utils.removeDir(dir_name)

def compress_zip(path,remove_source=False):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith('.zip'):
                full_path=root+'/'+f
                need_name=f.replace('.zip','')
                # print('need_name:%s' %need_name)
                utils.unzip(full_path)
                dir_path=full_path.replace('.zip','')
                rename_file(path,dir_path,need_name)
                if remove_source==True:
                    os.remove(full_path)

def get_file_number(path):
    for root,dirs,files in os.walk(path):
        return len(files)

def handle_many(start,end,use_cookies):

    for i in range(start,end):
        #合同协议
        menu_url='http://www.ppt118.com/words/list/t122/r1_{}.html'.format(i)

        detail_li=get_detail_li(menu_url)
        for title,download_url in detail_li:
            downloaded=utils.get_changed(com.path,need_ext=False)
            if not title in downloaded:
                download_path=com.path+'/'+title+'.zip'
                download_file(download_url,download_path,use_cookies=use_cookies)
    remove_empty_file(com.path)


def main_many(func,start,end):
    try:
        flag=True
        need_num=(end-start)*38
        count=2

        while flag:
            if count%2==0:
                func(start,end,use_cookies=False)
                time.sleep(2)
            else:
                func(start,end,use_cookies=True)


            count+=1
            if get_file_number(com.path)>=need_num:
                flag=False
    except Exception as e:
        print(e)
        func(start,end,func)




if __name__=='__main__':
    # for i in range(1,2):
    #     main(i)
    # remove_empty_file(com.path)
    # main_many(handle_many,2,3)

    compress_zip(com.path,remove_source=True)
