#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/12/7'

import os,re,time,shutil,random



import config as cfg
import common as com
import utils
from crawl_jianli import replace_text_jianli

test_path=com.path
save_path=cfg.save_path


def handle_word(path=test_path,root_path=test_path,remove_source=False):
    for root,dirs,files in os.walk(path):
        for f in files:
            if f.endswith(('.docx','.doc')):
                full_path=root+'/'+f
                file_size=os.path.getsize(full_path)
                if file_size>=10000:
                    filter_word_li=[' ']
                    for fw in filter_word_li:
                        if fw in f:
                            f=f.replace(fw, '')
                    final_path=root_path+'/'+f
                    if not os.path.exists(final_path):
                        shutil.move(full_path,final_path)

    if remove_source==True:
        li=[]
        for root,dirs,files in os.walk(path):
            for dir in dirs:
                print(dir)
                utils.removeDir(root+'/'+dir)
            for f in files:
                if not f.endswith(('.docx','.doc')):
                    fp=root+'/'+f
                    li.append(fp)

        for i in li:
            os.remove(i)

    print('简历处理完成！')

def replace_word(path=test_path):
    filter_li=[' ','+','&','副本','.']
    for i in filter_li:
        for root,dirs,files in os.walk(path):

            for dir in dirs:
                if i in dir:
                    dir_full_path=root+'/'+dir
                    new_dir_name=dir.replace(i,'')
                    new_dir_path=root+'/'+new_dir_name
                    try:
                        os.rename(dir_full_path,new_dir_path)
                    except Exception as e:
                        print(e)

            for f in files:
                file_name,ext=os.path.splitext(f)
                if i in file_name:
                    file_full_path=root+'/'+f
                    new_file_name=file_name.replace(i,'')
                    new_file_path=root+'/'+new_file_name+ext
                    try:
                        os.rename(file_full_path,new_file_path)
                    except Exception as e:
                        print(e)

    for root,dirs,files in os.walk(path):
        for f in files:
            if not '简历' in f:
                try:
                    file_name,ext=os.path.splitext(f)
                    file_full_path=root+'/'+f
                    find_num=re.findall('\d+$',file_name)
                    if find_num!=[]:
                        file_name=file_name.replace(find_num[0],'')
                        new_fn=file_name+'个人简历'+str(random.randint(1,1000))+ext
                    else:
                        new_fn=file_name+'个人简历'+ext
                    new_f_path=root+'/'+new_fn
                    os.rename(file_full_path,new_f_path)
                    print('{}:已修改'.format(f))
                except Exception as e:
                    print(e)
            res=re.findall('^(\d+).+?\.docx',f)
            if res!=[]:
                res=res[0]
                file_full_path=root+'/'+f
                new_file_name=f.replace(res,'')
                new_file_path=root+'/'+new_file_name
                if os.path.exists(new_file_path):
                    new_file_path=root+'/'+new_file_name.replace('.docx','')+str(random.randint(1,1000))+'.docx'

                try:
                    os.rename(file_full_path,new_file_path)
                except Exception as e:
                    print(e)

    print('多余字符串过滤完毕!')

def main():
    handle_word(test_path,remove_source=True)
    utils.doc_to_docx(test_path,del_origin_file=True)
    time.sleep(1)
    replace_word(test_path)
    replace_text_jianli(test_path)



if __name__=='__main__':

    # 记得先解压压缩包再运行
    main()


