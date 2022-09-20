#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/10/8'

import random

YEAR='2020'
MONTH='8月'
DATE='2020-8-4'

save_path='g:/资源/WORD资源/{}/{}/{}'.format(YEAR,MONTH,DATE)
# backup_path='g:/资源/WORD资源/{}/{}/{}/松鼠纯备份'.format(YEAR,MONTH,DATE)
pic_path='e:/图片保存/{}_WORD'.format(DATE)


word_filter=['精心整理','Welcome To','Download','!!!','欢迎您的下载，资料仅供参考！','仅供个人学习参考','---文章来源网络']
email_filter_one='@'


name_filter=['20{}{}'.format(m,n) for m in range(0,2) for n in range(0,10)]
name_filter=name_filter[:-2]
name_2_filter=[' ','超视立']
name_filter.extend(name_2_filter)


filter_after_word=['文档','Word文档','副本','范文','完整文档','参考']


ss_cookies={
    "Cookie":"lloij=0355b79de2fc4fafafeb8048dbe5a2a7; Hm_lvt_cf790a0e58ccc64ef6bfd1d01e8cb5d4=1570696137,1570701353,1570744383,1570816967; Hm_lpvt_cf790a0e58ccc64ef6bfd1d01e8cb5d4=1570818889; ASP.NET_SessionId=sm2jyry45nw3zwebu2q3pkvq; Hm_lvt_aa596a965d2070881ab5457040e61da3=1570696152,1570701360,1570744385,1570818893; Hm_lpvt_aa596a965d2070881ab5457040e61da3=1570819164"

}


fm_headers={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
    'Cookie':'showstl=yes; PHPSESSID=54gnel4rmn0ic6fvrhv5eqnco7; feemootoken=MDAwMDAwMDAwMJbfe2eakXHNtMuAro-Xs8-DdX3JmKqZnYPOq4eEqXatf7qKoJuKbaE.MDAwMDAwMDAwMJbfe6eYbH3NtMuAlp2ZxduRn4nNmLpvqoDSr6mTt3qnetOdn5l8eZqytouYh7qr34WsjNiEpq2djr3JqnrOoKx_upxqhH11mbPMd5OFqayXmImEyoXMf62BucysgKiYaHrUYXQ.d7268c28c9769fa77d9d2da672f35430cbdf780dfdca5cbec130d0839fa16e87'
}

# 11.4
# 'UM_distinctid=16d7cf651e6509-06124223a199e2-67e1b3f-1fa400-16d7cf651e7c4b; user_logo=http%3A%2F%2Fthirdqq.qlogo.cn%2Fg%3Fb%3Doidb%26k%3D36R8B8uLaTGxyJ52eficV7Q%26s%3D100%26t%3D1555697451; nickname=%E9%A3%8E; CNZZDATA1274869205=531170498-1572382964-null%7C1572382964; __root_domain_v=.ppt118.com; _qddaz=QD.87588.o8kpq4.k2cc7wvp; id=10315549; time=2019-10-30+05%3A02%3A38; key=4f773a99d6b81599d330ea600bfb4a9c; user_state_ppt=1; ci_session=9abrakhr6ah89d8vimu2f4g745h8j7dv; CNZZDATA1274869199=1308082960-1569758122-https%253A%252F%252Fwww.ppt118.com%252F%7C1572869929; _qdda=3-1.3p0sea; _qddab=3-on659e.k2kfau8t; _qddamta_2852163597=3-0; CNZZDATA1274869202=1980686756-1572378490-null%7C1572870073',


fengyun_cookies={
    'Cookie':'UM_distinctid=16d7cf651e6509-06124223a199e2-67e1b3f-1fa400-16d7cf651e7c4b; user_logo=http%3A%2F%2Fthirdqq.qlogo.cn%2Fg%3Fb%3Doidb%26k%3D36R8B8uLaTGxyJ52eficV7Q%26s%3D100%26t%3D1555697451; nickname=%E9%A3%8E; CNZZDATA1274869205=531170498-1572382964-null%7C1572382964; __root_domain_v=.ppt118.com; _qddaz=QD.87588.o8kpq4.k2cc7wvp; id=10315549; time=2019-10-30+05%3A02%3A38; key=4f773a99d6b81599d330ea600bfb4a9c; user_state_ppt=1; CNZZDATA1274869199=1308082960-1569758122-https%253A%252F%252Fwww.ppt118.com%252F%7C1572875338; _qddamta_2852163597=3-0; CNZZDATA1274869202=1980686756-1572378490-null%7C1572875473; ci_session=3deatulq3ksbgoc21vd6pnt8jm714mq2; _qdda=3-1.48xr6a; _qddab=3-ckrysk.k2kjp0g6'
}
JOB_TYPE=['简历','求职']

USEFUL_TYPE=['合同','协议','进货','出货','采购','策划','计划','报告']

COMMON_TYPE=['稿','公文','写作','致辞','党','函','日记','周记','手册','个人总结']

PERSON_TYPE=['表','制度','审批','登记','申请','证明','细则',
             '入职','离职','试用','请假','转正','人事']


# pdf_handle config
max_worker=5

common_headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.87 Safari/537.36',
}

def get_headers():
    headers = {
        'User-Agent' : 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:46.0) Gecko/20100101 Firefox/{}.0'.format(random.randint(6000,65000)),

    }
    return headers

if __name__=='__main__':
    print(name_filter)
