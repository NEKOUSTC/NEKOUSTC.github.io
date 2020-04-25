#!/usr/bin/python
# -*- coding: UTF-8 -*-
import json  #JSON 编码和解码器
import xlrd  #表格读取
import xlwt  #表格写入（虽然没见用上）
import argparse  #命令行选项、参数和子命令解析器
import sys  #系统相关的参数和函数
import os  #多种操作系统接口
import time  #时间
import urllib.parse  #用于解析 URL
import urllib  #URL 处理模块
import fnmatch  #文件名匹配
import shutil  #高级文件操作

today = time.strftime("%Y-%m-%d", time.localtime())  #时间转换为YYYY-MM-DD形式

workbook = xlrd.open_workbook('猫咪档案0425.xlsx')  #打开猫咪档案的表格文档

data = workbook.sheets()[0]  #建立工作表的一个用于数据操作的副本

rowNum = data.nrows  # sheet行数
colNum = data.ncols  # sheet列数

# 获取所有单元格的内容
data_list = []  #建立数据列表
for i in range(12, rowNum):  #从第13行开始到有内容的末尾
    rowlist = []  #建立行列表
    for j in range(colNum):  #从第一列到最后
        ctype = data.cell(i, j).ctype  #获取单元格数据类型
        cell = data.cell_value(i, j)  #获取单元格数据的值
        if data.cell(i, j).ctype == 3:  #如果类型为3（应该是日期DATE）
            dt = xlrd.xldate.xldate_as_tuple(data.cell_value(i, j), 0)  #获取日期值
            rowlist.append('%04d-%02d-%02d' % dt[0:3])  #在行列表末尾添加日期元素
            continue  #转到下一个单元格
        if ctype == 2 and cell % 1 == 0.0:  # ctype为2且为浮点  #如果类型为2（应该是数值）
            cell = int(cell)  #转换为整数
            rowlist.append(cell)  #附加到行列表末尾
            continue  #转到下一个单元格
        rowlist.append(data.cell_value(i, j))  #如果不是上面的两种类型，就直接附加在行列表末尾
    data_list.append(rowlist)  #建立完一个行列表后，将其附加在数据列表末尾，并开始下一行

# 输出所有单元格的内容

rowNum -= 12  #去掉前面用于描述的12行

"""
#这部分用于看输入的数据列表是不是正确
for i in range(rowNum):
    for j in range(colNum):
        print(data_list[i][j],end=" ")  #原本是(data_list[i][j],' ',end="")，试试看这样行不行？
    print('\n')
"""

"""
新版项目：0、；1、；2、名字；3、是否写入图鉴；4、昵称；5、毛色（文字描述）；*6、毛色分类（序号；7、出没地点
8、性别；9、状况；10、绝育情况；11、绝育时间；12、出生时间；13、外貌；14、性格；
15、第一次被目击时间；//16、第一次被目击地点；17、关系；//18、备注；//19、线路；//20、送养时间；
21、；22、是否加音频；
"""
#下面是一个好大的列表（lambda是建立一个函数，输入在:前，输出在:后，同时输出要先判断if再赋值）


labels = [
    [2, '名字', lambda x:'【还没有名字】' if len(x) < 1 else x],
    [3, '是否写入图鉴', lambda x:x],
    [4, '昵称', lambda x:x],
    #[5, '毛色', lambda x:x],
    [6, '毛序', lambda x: '纯色' if x == 5 else '玳瑁及三花' if x == 4 else '奶牛' if x == 3 else '橘猫及橘白' if x == 2 else '狸花' if x == 1 else x if x == 0 else '' ],
    [8, '性别', lambda x:'公' if x == 1 else '母' if x == 0 else '未知'],
    [9, '状况', lambda x:'不明' if len(x) < 1 else x],
    [10, '绝育情况', lambda x:'已绝育' if x == 1 else '未绝育' if x == 0 else '未知/可能不适宜绝育'],
    [11, '绝育时间', lambda x:str(x)],
    [12, '年龄', lambda x:x],
    [13, '外貌', lambda x:x],
    [14, '性格', lambda x:x], #'亲人可抱' if x == 6 else '亲人不可抱 可摸' if x == 5 else '薛定谔亲人' if x == 4 else '吃东西时可以一直摸' if x == 3 else '吃东西时可以摸一下' if x == 2 else '怕人 安全距离1m以内' if x == 1 else '怕人 安全距离1m以外' if x == 0 else '未知 数据缺失' ],
    [15, '第一次被目击时间', lambda x: str(x)],
    [17, '关系', lambda x: str(x)],
    [22, '是否加音频', lambda x:x]
]

data_json = []  #建立一个数据列表
#wnls=os.listdir('文案')  #建立一个文案文件夹里文件的名称的列表

for i in range(rowNum):  #遍历全部行
    json_line = {}  #定义字典json_line

    if len(data_list[i][5]) < 1: # 毛色都不知道那就是没有【x
        continue  #没有就直接下一行【x

    for j in labels:  #对labels这个列表中的元素挨个赋值
        json_line[j[1]] = j[2](data_list[i][j[0]])  #关键字：j[1]这列（即labels的第二列）；值：j[2]列，输入的参数为data_list第i行对应j[0]所代表的序号的值

    data_json.append(json_line)  #将这一行赋值后的字典附加到数据列表后
    #print(json_line)  #打印这个列表以便检查

#创建文章文件夹
if not os.path.exists('_posts'):  #如果本路径不存在'_posts'文件夹
    os.makedirs('_posts')   #建立'_posts'文件夹

# 用于搜索关系链接
names = []  #建立names列表
for line in data_json:  #遍历data_json这个数据列表的每一行（每一行都是一个字典）
    if line['是否写入图鉴'] != '':  #如果写入图鉴了
        names.append(line['名字'])  #就把名字附加在names列表后
#print(names)

for line in data_json:  ##遍历data_json这个数据列表的每一行（每一行都是一个字典）
    if line['是否写入图鉴'] != '':  #如果写入图鉴了就继续
        for f in os.listdir('_posts'):
            if fnmatch.fnmatch(f,'[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]-'+ line['名字'] +'.md'):
                shutil.move('_posts/' + f,'_backup/' + f,)
        
        with open('_posts/'+ today + '-' + line['名字'] + '.md', 'w') as f:  #就建立名字.md并打开
            f.write( '---\n' 'title: ' +line['名字'] + '\n' 'tags: ')  #这里对照现有文档好好改改吧
            for j in labels:  #遍历labels列表中的各项
                if j[1] == '状况' or j[1] == '毛序':
                    f.write(line[j[1]] + ' ')
            f.write( '\n---\n\n')
            for j in labels:  #遍历labels列表中的各项
                if j[1] == '状况' or j[1] == '毛序':
                    continue
                if j[0] == 22:  #如果第一列的序号为22（音频）则跳过后面
                    continue
                if str(line[j[1]]) == '' or j[1] == '是否写入图鉴':  #如果值是空的或者是否写入图鉴，均跳过，并开始下一行（因为不需要告诉大家有没有写入图鉴
                    continue
                f.write('**' + j[1] + '**:' + str(line[j[1]]) + '\n\n')
             #编写日期
            f.write('*编写日期:' + today + '*\n\n')  #写入编写日期（即刚开始时获取的日期字符串）
            '''#增加关系跳转项
            f.write('关系 : ')  #
            for i in names:
                if i in line['关系'] and i!=line['名字']:
                    f.write( '[' + i + '](https://NEKOUSTC.github.io/'+ time.strftime("%Y/%m/%d/", time.localtime()) + line['名字'] +')；')       
            f.write( '\n')
            '''
            
            # 后面的图片数
            if line['是否写入图鉴'] != 0:
                f.write( '**附录图片**：\n\n')
                for i in range(int(line['是否写入图鉴'])):  #在表格中的这项，数字代表图片数
                    f.write('!['+ line['名字'] +'{}'.format(i+1)+'](http://q9a0pgz83.bkt.clouddn.com/cats/m_'+ line['名字'] +'{}'.format(i+1)+'.jpg)    \n')  #将format(i+1)赋到{}的位置 （format：格式化赋值？）下面音频的也差不多
                    f.write( '[查看原图](http://q9a0pgz83.bkt.clouddn.com/cats/l_'+ line['名字'] +'{}'.format(i+1)+'.jpg)    \n')
            
            '''
            # 后面的音频数
            if line['是否加音频']:
                audio = '//pku-lostangel.oss-cn-beijing.aliyuncs.com/' + line['名字']  #音频来自阿里云主机？
                audio = urllib.parse.quote(audio)  #url解析并赋值回来
                f.write('audioArr: [\n')
                for i in range(line['是否加音频']):  #在表格中的这项，数字代表音频数
                    f.write( '{\n ' + "src: 'https:" + audio + "{}.m4a'".format(i+1) + ',\nbl: false\n},\n')
                f.write("],\n  audKey: '', \n},\n")
            else:
                f.write('},')
            with open('js.txt','r') as f2:
                f.write(f2.read())  #将js.txt中的内容接着写入js文档（主要内容：音频播放相关控件，刷新、转发等控件）
'''        

'''
#几个分页的index内容（显示哪些猫）
health = []  #健康
fostered = []  #送养
dead = []  #离世
unknown = []  #不明
nainiu = []  #奶牛
sanhua = []  #玳瑁及三花
chunse = []  #纯色
lihua = []  #狸花
ju = []  #橘猫及橘白
suoyou = []  #所有
'''

"""
# 分类
for i in range(rowNum):
    if data_list[i][3] != '':  #这是最初的数据列表，第[3]列为是否写入图鉴
        if data_list[i][9] == '离世':  #第[9]列为状况
            dead.append(data_list[i][2])  #在离世列表中附加名字
        if data_list[i][9] == '送养':
            fostered.append(data_list[i][2])  #在送养列表中附加名字
        if (data_list[i][9] == '不明' or data_list[i][9] == '许久未见'or data_list[i][9] == '失踪'):#这倒是个单独加猫的办法→ and data_list[i][2] != '花灵灵':  #在不明列表中附加名字
            unknown.append(data_list[i][2])  #在不明列表中附加名字
        if (data_list[i][9] == '健康' or data_list[i][9] == '口炎'):# and data_list[i][2] != '出竹':  #健康列表？没有直接弄啊。。。
            if data_list[i][6] == 1:  #这里开始分花色了，但是上面那几种就不分花色了吗？
                lihua.append(data_list[i][2])
            if data_list[i][6] == 2:
                ju.append(data_list[i][2])
            if data_list[i][6] == 3:
                nainiu.append(data_list[i][2])
            if data_list[i][6] == 4:
                sanhua.append(data_list[i][2])
            if data_list[i][6] == 5:
                chunse.append(data_list[i][2])

#lihua.insert(0,'出竹')
#unknown.insert(0,'花灵灵')
health = lihua + ju + nainiu + sanhua + chunse  #各种颜色加起来就是健康的全部
suoyou = health + unknown + dead  #健康+不明+离世就是所有（不算送养吗。。。好吧确实送养不算校内了
"""
