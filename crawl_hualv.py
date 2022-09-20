#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/29'

from gevent import monkey
monkey.patch_all()
import os,time,re,shutil
from threading import Thread

import requests
from pyquery import PyQuery as pq


import config as cfg
import common as com
import utils
import gevent

test_url='https://www.66law.cn/contractmodel/37465.aspx'

base_url='https://www.66law.cn'

menu_url='{}/contractmodel/all/page_1.aspx'.format(base_url)

detail_url='https://www.66law.cn/contractmodel/37465.aspx'

title_filter_str='一二三四五六七八九'
title_filter_li=['({})'.format(i) for i in title_filter_str]
title_filter_li2=['（{}）'.format(i) for i in title_filter_str]

title_filter_li.extend(title_filter_li2)


def get_detial_li(url):
    li=[]
    res=requests.get(url,headers=cfg.get_headers()).content.decode('utf-8')
    html=pq(res)
    datas=html('table.pact-list div.tit')

    for data in datas.items():
        title=data('a').text()
        url=data('a').attr('href')
        url=base_url+url
        print(title)

        if '[' in title:
            title=title.replace('[','(')

        if ']' in title:
            title=title.replace(']',')')

        for i in title_filter_li:
            if i in title:
                title=title.replace(i,'')

        filter_words=re.findall('\(\d+\)$|（\d+）$|\d+$',title)
        if filter_words!=[]:
            for fw in filter_words:
                title=title.replace(fw,'')


        print(url)
        print('-'*30)
        li.append((title,url))
    return li

def get_content(url):
    res=requests.get(url,headers=cfg.get_headers()).content.decode('utf-8')
    html=pq(res)


    content=html('div.det-nr').html()
    content=str(content)

    p_ele=re.findall('<p.*?>',content)
    content=content.replace('</p>','\n')

    div_ele=re.findall('<div.*?>',content)
    content=content.replace('</div>','\n')

    style_ele=re.findall(' style=".+?"|<style.+?</style>',content,re.S)



    a_ele=re.findall('<a.*?>',content)
    content=content.replace('</a>','')

    span_ele=re.findall('<span.*?>',content,re.S)
    content=content.replace('</span>','')

    form_ele=re.findall('<form.*?>.+?</form>',content)


    content=content.replace('<br/>','\n')

    special_sign=re.findall('&.+?;',content)

    font_ele=re.findall('<font.*?>',content)
    content=content.replace('</font>','')

    # print(a_ele)
    if p_ele!=[]:
        for p in p_ele:
            content=content.replace(p,'\n')

    if div_ele!=[]:
        for div in div_ele:
            content=content.replace(div,'\n')

    if style_ele!=[]:
        for style in style_ele:
            content=content.replace(style,'')

    if a_ele!=[]:
        for a in a_ele:
            content=content.replace(a,'')

    if span_ele!=[]:
        for sp in span_ele:
            content=content.replace(sp,'')

    if form_ele!=[]:
        for fm in form_ele:
            content=content.replace(fm,'')

    if special_sign!=[]:
        for s in special_sign:
            content=content.replace(s,'')

    if font_ele!=[]:
        for ft in font_ele:
            content=content.replace(ft,'')

    filter_word=['?','13;']
    for w in filter_word:
        if w in content:
            content=content.replace(w,' ')


    return content


    # content.replace('<p>','')
    # content.replace('</p>','')


def handle_process(title,detail_url):
    try:
        content=get_content(detail_url)
        file_name=title+'.doc'
        content_title='<h1>{}</h1>\n\n'.format(title)
        downloaded=utils.get_changed(com.path,need_ext=False)
        if title not in downloaded:
            content=content_title+content
            com.save_file(com.path,file_name,content)
    except Exception as e:
        print(e)


def choice_cls(path):
    need_word_li=['房','劳动','股','采购','购','借','贷','知识',
                  '产权','保险','证券','委托','投资','商标','专利',
                  '买卖','销售','租','保密','版权','经营','转让','销',
                  '代理','饭店','酒店','聘','制度','信托','期货','期权',
                  ]
    need_word_li=sorted(need_word_li,key=lambda x:len(x),reverse=True)

    for root,dirs,files in os.walk(path):
        for f in files:
            for i in need_word_li:
                if i in f:
                    try:
                        full_path=root+'/'+f
                        dst_path='g:/资源/WORD资源/2019/{}/{}'.format(cfg.MONTH,f)
                        shutil.move(full_path,dst_path)
                    except Exception as e:
                        print(e)

    print('过滤转移完毕！')


def main(num):
    # 全部类别
    menu_url='https://www.66law.cn/contractmodel/all_1/page_{}.aspx'.format(num)

    detail_li=get_detial_li(menu_url)
    # for title,detail_url in detail_li:
    gevent.joinall([
        gevent.spawn(handle_process,title,detail_url) for title,detail_url in detail_li
    ])



if __name__=='__main__':

    # 到11月8号
    # for i in range(600,761):
    #     # main(i)
    #     t=Thread(target=main,args=(i,))
    #     t.start()



    # utils.rename_file(com.path,cfg.word_filter)
    # utils.doc_to_docx(com.path,del_origin_file=True)
    # com.modify_docx(com.path)
    # utils.rename_file_front(com.path)
    # utils.rename_file_end(com.path)


    path='g:/资源/WORD资源/2019/11月/可用'
    choice_cls(path)
