#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/29'

import os

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import config as cfg

path='g:/资源/WORD资源/{}/{}/测试'.format(cfg.YEAR,cfg.MONTH)



def save_file(path,file_name,content):
    full_path=path+'/'+file_name
    with open(full_path,'w',encoding='utf8') as f:

        if isinstance(content,list):
            for text in content:
                # print(text)
                f.write(text)
        else:
            f.write(content)

    print('{}:写入成功'.format(file_name))


def modify_docx(path):

    for root,dirs,files in os.walk(path):
        for f in files:
            try:
                if f.endswith('docx'):
                    full_path=root+'/'+f
                    doc=Document(full_path)
                    for para in doc.paragraphs:
                        if '<h1>' in para.text:
                            para.paragraph_format.alignment=1
                            print('{}:{}已居中显示'.format(f,'h1'))

                            for run in para.runs:
                                replace_h1(run)
                                print('{}:{}字体已格式化'.format(f,'h1'))
                            # replace_h1(run)

                        else:
                            paragraph_format = para.paragraph_format
                            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


                        if 'strong' in para.text:
                            r=None
                            for run in para.runs:

                                r=replace_strong(run,r)
                                print('{}:{}字体已加粗'.format(f,'strong'))

                        if 'span' in para.text:
                            for run in para.runs:
                                replace_span(run)
                                print('{}:{}已去除'.format(f,'span'))


                    doc.save(full_path)

            except Exception as e:
                print(e)

def replace_h1(run):
    run.font.size=300000
    run.font.bold=True
    run.text=run.text.replace('<h1>','')
    run.text=run.text.replace('</h1>','\n\n')

def replace_span(run):

    run.text=run.text.replace('<span>','')
    run.text=run.text.replace('</span>','')


def replace_strong(run,last_run):
    run.font.bold=True
    # print(run.text)

    run.text=run.text.replace('<strong>','')
    run.text=run.text.replace('</strong>','')

    if last_run !=None:
        sum_run=last_run.text+run.text
        if sum_run=='<strong>' or sum_run=='</strong>':
            run.text=''
            last_run.text=''

    return run



def choice_text_to_replace(text):
    len_text=len(text)
    if len_text<=3:
        word='X'*len_text
    else:
        word='X'*(len_text-1)

    return word




