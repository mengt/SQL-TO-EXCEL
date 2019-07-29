#!/usr/bin/env python
#-*- coding:utf-8-*-

import io
import xlwt
import xlrd
import openpyxl
import re
import sys
import codecs
import httplib
import md5
import urllib
import random
import json
import time
#https://www.cnblogs.com/hushaojun/p/7792550.html
# Text values for colour indices. "grey" is a synonym of "gray".
# The names are those given by Microsoft Excel 2003 to the colours
# in the default palette. There is no great correspondence with
# any W3C name-to-RGB mapping.
_colour_map_text = """\
    aqua 0x31
    black 0x08
    blue 0x0C
    blue_gray 0x36
    bright_green 0x0B
    brown 0x3C
    coral 0x1D
    cyan_ega 0x0F
    dark_blue 0x12
    dark_blue_ega 0x12
    dark_green 0x3A
    dark_green_ega 0x11
    dark_purple 0x1C
    dark_red 0x10
    dark_red_ega 0x10
    dark_teal 0x38
    dark_yellow 0x13
    gold 0x33
    gray_ega 0x17
    gray25 0x16
    gray40 0x37
    gray50 0x17
    gray80 0x3F
    green 0x11
    ice_blue 0x1F
    indigo 0x3E
    ivory 0x1A
    lavender 0x2E
    light_blue 0x30
    light_green 0x2A
    light_orange 0x34
    light_turquoise 0x29
    light_yellow 0x2B
    lime 0x32
    magenta_ega 0x0E
    ocean_blue 0x1E
    olive_ega 0x13
    olive_green 0x3B
    orange 0x35
    pale_blue 0x2C
    periwinkle 0x18
    pink 0x0E
    plum 0x3D
    purple_ega 0x14
    red 0x0A
    rose 0x2D
    sea_green 0x39
    silver_ega 0x16
    sky_blue 0x28
    tan 0x2F
    teal 0x15
    teal_ega 0x15
    turquoise 0x0F
    violet 0x14
    white 0x09
    yellow 0x0D"""
 
pattern="[A-Z]"
reload(sys)
sys.setdefaultencoding('utf8')
appKey = '4330c2608a6d4b4b'
secretKey = 'YaBZTivi8rZ5JGeYPs5isevmXygdY1El'

def sendDate(q):
    httpClient = None
    myurl = '/api'
    #myurl = 'openapi.do'
    q = q
    fromLang = 'EN'
    toLang = 'zh-CHS'
    salt = random.randint(1, 65536)

    sign = appKey+q+str(salt)+secretKey
    m1 = md5.new()
    m1.update(sign)
    sign = m1.hexdigest()
    myurl = myurl+'?appKey='+appKey+'&q='+urllib.quote(q)+'&from='+fromLang+'&to='+toLang+'&salt='+str(salt)+'&sign='+sign
    #myurl = myurl+'?keyfrom='+appKey+'&key='+secretKey+'&type=data&doctype=json&version=1.1&q='+q
    try:
        httpClient = httplib.HTTPConnection('openapi.youdao.com')
        #httpClient = httplib.HTTPConnection('fanyi.youdao.com')
        httpClient.request('GET', myurl)

        #response是HTTPResponse对象
        response = httpClient.getresponse()
        rejson = response.read()

        s = json.loads(rejson)
        return s['translation'][0]
    except Exception as e:
        print(e)
    finally:
        if httpClient:
            httpClient.close()

# 
#   `id` int(10) unsigned（无符号（unsigned）和有符号（signed）） NOT NULL AUTO_INCREMENT（自动生成）,
#   `id_alert` varchar(36) NOT NULL,
#   `id_action_plan` int(10) unsigned NOT NULL,
#   `created_by` int(10) unsigned NOT NULL,
#   `version_c` int(11) NOT NULL DEFAULT '0',
#   PRIMARY KEY (`id`,`id_alert`,`id_action_plan`), 组合主键
#   UNIQUE KEY `id_alert` (`id_alert`,`id_action_plan`),  唯一性约束
#   KEY `fk_action_plan_alert_trigger_action_plan` (`id_action_plan`), 索引
#   KEY `fk_action_plan_alert_trigger_user` (`created_by`), 索引
#   CONSTRAINT `fk_action_plan_alert_trigger_action_plan` FOREIGN KEY (`id_action_plan`) REFERENCES `action_plan` (`id`) ON DELETE CASCADE,
#   CONSTRAINT `fk_action_plan_alert_trigger_alert` FOREIGN KEY (`id_alert`) REFERENCES `alert` (`uuid`) ON DELETE CASCADE,
#   CONSTRAINT `fk_action_plan_alert_trigger_user` FOREIGN KEY (`created_by`) REFERENCES `user` (`idUser`) ON DELETE CASCADE
#   外键        外键名字                           表内字段名字               关联      外键表名   外键字段名   联级删除
# ) ENGINE=InnoDB DEFAULT CHARSET=utf8;

#python操作excel主要用到xlrd和xlwt这两个库，即xlrd是读excel，xlwt是写excel的库。
#由于xlrd不能对已存在的xlsx文件，进行修改！所以必须使用OpenPyXL

#设置表格样式
def set_style(name,height,bold=False,pattern_fore_colour='null',borders = False):
    style = xlwt.XFStyle()
    #字体样式
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    #背景色
    if pattern_fore_colour != "null":
        background = xlwt.Pattern()
        background.pattern = xlwt.Pattern.SOLID_PATTERN
        background.pattern_fore_colour = xlwt.Style.colour_map[pattern_fore_colour]
        ## May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on.
        style.pattern = background
    if borders:
        #边框
        # borders.left = xlwt.Borders.THIN
        # NO_LINE： 官方代码中NO_LINE所表示的值为0，没有边框
        # THIN： 官方代码中THIN所表示的值为1，边框为实线
        borders = xlwt.Borders()
        color = xlwt.Style.colour_map["black"]
        borders.left = color
        borders.left = xlwt.Borders.THIN
        borders.right = color
        borders.right = xlwt.Borders.THIN
        borders.top = color
        borders.top = xlwt.Borders.THIN
        borders.bottom = color
        borders.bottom = xlwt.Borders.THIN
        # 定义格式
        style.borders = borders
    # alignment = xlwt.Alignment() # Create Alignment
    # alignment.horz = xlwt.Alignment.HORZ_LEFT  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    # alignment.vert = xlwt.Alignment.VERT_TOP   # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    # alignment.wrap = 1  #自动换行
    # style.alignment = alignment
    return style

def update_style(pattern_fore_colour):
    style = xlwt.XFStyle()
    background = xlwt.Pattern()
    background.pattern = xlwt.Pattern.SOLID_PATTERN
    background.pattern_fore_colour = xlwt.Style.colour_map[pattern_fore_colour]
    ## May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on.
    style.pattern = background
    return style

#写Excel
def write_excel():

    excle_open = xlwt.Workbook(encoding = 'utf-8') 
    sheet1 = excle_open.add_sheet('sql_mode',cell_overwrite_ok=True)
    row0 = ["字段","主键","外键","索引","数据类型","符号类型","非空","默认值","唯一性约束","描述"]
    #普通主题
    default = set_style('Times New Roman', 220, True)
    #table主题
    default_table = set_style('Times New Roman', 220, True, 'null', True)
    #row0主题
    default_colo = set_style('Times New Roman', 220, True,'yellow', True)
    #主键背景色
    PRIMARY_KEY_colo = update_style('orange')
    #外键背景色
    CONSTRAINT_colo = set_style('Times New Roman', 220, True,'red', False)
    #索引背景色
    KEY_colo = update_style('green')

    #控制列数
    x_num = 0
    #控制行数
    y_num = 0
    #控制跳过被注释的行
    pass_loop = False

    #table内循环范围,由字典控制
    y_num_for_field = {}
    #`objectName`(255)) 首先在sql文件中讲(`objectName`(255))这个改成(`objectName`）
    with io.open('kinton-schema.sql','r', encoding='utf-8') as sql_mode:
        sql_list = sql_mode.readlines()
    for i in range(0,len(sql_list)):
        #进行筛选
        #控制跳过被注释的行
        if sql_list[i].startswith("/*"): 
            pass_loop = True
        if sql_list[i].endswith("*/;\n"):
            pass_loop = False
        if pass_loop:
            continue
        #录入行
        if sql_list[i].startswith("CREATE TABLE"):
            y_num_for_field.clear()  
            y_num += 1
            sheet1.write(y_num, 0, "表名：", default_table)
            sheet1.col(0).width = 6666
            #合并行单元格
            wait_string = sql_list[i].split("`")[1].split("`")[0]
            translation_string = sendDate(wait_string.replace('_',' '))
            time.sleep(0.2)
            print translation_string
            sheet1.write_merge(y_num,y_num,1,9,unicode(wait_string+":"+translation_string),default_table) 
            #sheet1.write(y_num, 1,sql_list[i].split("`")[1].split("`")[0] , default)
            
            #写第一行
            y_num += 1
            for i in range(0,len(row0)):
                sheet1.write(y_num,i,row0[i],default_colo)
            y_num += 1 

        #数据表字段
        sql_row = sql_list[i].strip()
        if sql_row.startswith("`"):
            #print sql_list[i]
            #["字段","主键","外键","索引","数据类型","符号类型","非空","默认值","唯一性约束"，"描述"]
            field_value = sql_row.split("`")[1].split("`")[0]
            sheet1.write(y_num, 0, field_value, default)
            if field_value.replace('_',' ').isupper():
                translation_field_value = sendDate(field_value)
            else:
                xxx= field_value.replace('_',' ')
                translation_field_value = sendDate(re.sub(pattern,lambda xxx:" "+xxx.group(0),xxx))   
            time.sleep(0.2)
            print translation_field_value
            sheet1.write(y_num, 9, unicode(translation_field_value), default)
            #加入到y_num_for_field字典中
            y_num_for_field[field_value] = y_num
            row_sql_list = sql_row.split(" ")
            #数据类型
            #row_sql_list[1] 
            sheet1.write(y_num, 4, row_sql_list[1], default)
            if "AUTO_INCREMENT" in row_sql_list:
                sheet1.write(y_num, 1, "AUTO_INCREMENT", default)
                y_num_for_field[y_num] = "AUTO_INCREMENT"
            #无符号（unsigned）和有符号（signed）
            if "unsigned"  in row_sql_list:
                sheet1.write(y_num, 5, u"无符号", default)
            if "signed"  in row_sql_list:
                sheet1.write(y_num, 5, u"有符号", default)
            if "NOT" in row_sql_list and "NULL" in row_sql_list:
                sheet1.write(y_num, 6, "True", default)
            if "DEFAULT" in row_sql_list:
                def_value = row_sql_list[row_sql_list.index('DEFAULT')+1]
                if "0000-00-00" in def_value:
                    def_value = def_value+" 00:00:00'"
                else:
                    def_value = def_value[:-1]
                sheet1.write(y_num, 7, def_value, default)
            y_num += 1

        #标识主键
        if sql_row.startswith("PRIMARY KEY"):
            field_value = sql_row.split("(`")[1].split("`)")[0].split("`,`")
            for i in field_value:
                row_num = y_num_for_field[i]
                if row_num in y_num_for_field:
                    sheet1.write(row_num, 1, y_num_for_field[row_num], PRIMARY_KEY_colo)
                else:
                    sheet1.write(row_num, 1, style = PRIMARY_KEY_colo)
        #标识唯一性约束
        if sql_row.startswith("UNIQUE KEY"):
            field_value = sql_row.split("(`")[1].split("`)")[0].split("`,`")
            for i in field_value:
                row_num = y_num_for_field[i]
                sheet1.write(row_num, 8, i, default)
        #表示索引
        if sql_row.startswith("KEY"):
            field_value = sql_row.split("(`")[1].split("`)")[0].split("`,`")
            for i in field_value:
                row_num = y_num_for_field[i]
                sheet1.write(row_num, 3, style = KEY_colo)
        #表示外键
        if sql_row.startswith("CONSTRAINT"):
            field_value = sql_row.split("FOREIGN KEY (`")[1].split("`) REFERENCES")[0].split("`,`")
            for i in field_value:
                row_num = y_num_for_field[i]
                sheet1.write(row_num, 2, sql_row.split("REFERENCES ")[1].split("ON DELETE CASCADE")[0], style = CONSTRAINT_colo)

    excle_open.save('test-p.xls')


if __name__ == '__main__':
    write_excel()
