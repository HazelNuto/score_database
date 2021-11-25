# -*- coding: UTF-8 -*-
import pandas as pd
import numpy as np
import math
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import xlrd
from pandas import DataFrame
import math
import random
import os
import json
import sys
import easygui as g

dataPath = "./data/result.xls"
o = dataPath
wb = xlrd.open_workbook(dataPath)
sheets = wb.sheet_names()
color_sch = [["#DC143C","#FFA500","#FFFF00","#3CB371","#00BFFF"],["#FFCDB2","#FFB4A2","#E5989B","#b5838D","#6D6875"],["#d8dbe2","#a9bcd0","#58a4b0","#b8d8ba","#5e6472"]]
color_switch = 2 #更换配色方案： 0炫彩色， 1淡雅暖色,  2淡雅冷色

df = pd.read_excel(dataPath, skiprows=0,keep_default_na=False)
nrows = df.shape[0]
ncols = df.shape[1]

def debug_print_data(data, class_cnt):
    for i in range(class_cnt):
        for j in range(5):
            print(data[i][j], end = ' ')
        print()

def debug_print_level(da, db, dc, dd, de, class_cnt):
    for i in range(class_cnt):
        print("%.2f %.2f %.2f %.2f %.2f" % (da[i], db[i], dc[i], dd[i], de[i]))

def update_paint(rate, addsum, cnt):
    for i in range(cnt):
        rate[i] += addsum[i]

while(True):

    choices = ["分析年级等级分布","分析班级等级分布","班级与年纪平均分差","设置","退出"]
    input_type = g.buttonbox(msg="操作类型", title="成绩整理与分析系统", choices=("分析年级等级分布","分析班级等级分布","班级与年纪平均分差","设置","退出"))

    
    if input_type == choices[0]:
        input_grade = int(g.buttonbox(msg="选择年级",title="成绩整理与分析系统",choices=("1","2","3","4","5","6")))
        input_sub = g.buttonbox(msg="选择学科",title="成绩整理与分析系统",choices=("语文","数学","英语","科学"))
        start_col = 0
        if input_sub == "语文":
            start_col = 9
        elif input_sub == "数学":
            start_col = 19
        elif input_sub == "英语":
            start_col = 27
        elif input_sub == "科学":
            start_col = 35
        else:
            g.msgbox("输入有误，请重试")
            continue
        data = []
        class_lst = []
        a_rate = []
        b_rate = []
        c_rate = []
        d_rate = []
        e_rate = []
        for i in range(nrows):
            if df.iloc[i][0] == '' or df.iloc[i][0] == '/':
                continue
            first_str = str(df.iloc[i][0])[0]
            if first_str >= '0' and first_str <= '9': #处理班级或年级
                grade = int(first_str)
                if grade == input_grade:
                    if str(df.iloc[i][0])[1] < '0' or str(df.iloc[i][0])[1] > '9': #年级
                        continue
                    class_lst.append(str(df.iloc[i][0]))
                    tup = []
                    for j in range(5):
                        rate_str = str(df.iloc[i,j+start_col])
                        rate_len = len(rate_str)
                        tup.append(float(rate_str[0:rate_len-1]))
                    data.append(tup)
                elif grade > input_grade:
                    break
        #data = pd.DataFrame(data=data, columns=['A等率','B等率','C等率','D等率','E等率'],index=class_lst)

        class_cnt = len(class_lst)
        if class_cnt == 0:
            g.msgbox("输入有误，请重试")
            continue

        for i in range(class_cnt):
            a_rate.append(data[i][0])
            b_rate.append(data[i][1])
            c_rate.append(data[i][2])
            d_rate.append(data[i][3])
            e_rate.append(data[i][4])

        mpl.rcParams["font.sans-serif"] = ["SimHei"]
        mpl.rcParams["axes.unicode_minus"] = False

        x = [i+1 for i in range(class_cnt)]
        #debug_print_data(data, class_cnt)
        #debug_print_level(a_rate, b_rate, c_rate, d_rate, e_rate, class_cnt)

        paint_rate = [0.0 for i in range(class_cnt)]
        plt.bar(x, a_rate, align="center", color=color_sch[color_switch][0], tick_label=class_lst, label="A等率")
        update_paint(paint_rate, a_rate, class_cnt)
        plt.bar(x, b_rate, align="center", bottom=paint_rate, color=color_sch[color_switch][1], label="B等率")
        update_paint(paint_rate, b_rate, class_cnt)
        plt.bar(x, c_rate, align="center", bottom=paint_rate, color=color_sch[color_switch][2], label="C等率")
        update_paint(paint_rate, c_rate, class_cnt)
        plt.bar(x, d_rate, align="center", bottom=paint_rate, color=color_sch[color_switch][3], label="D等率")
        update_paint(paint_rate, d_rate, class_cnt)
        plt.bar(x, e_rate, align="center", bottom=paint_rate, color=color_sch[color_switch][4], label="E等率")


        plt.xlabel("班级")
        plt.ylabel("比例")

        plt.legend()

        plt.show()

    elif input_type == choices[1]:
        input_class = g.enterbox(msg="输入分析班级", title="成绩整理与分析系统", default="三位数字",)
        input_sub = g.buttonbox(msg="选择学科",title="成绩整理与分析系统",choices=("语文","数学","英语","科学"))
        start_col = 0
        if input_sub == "语文":
            start_col = 9
        elif input_sub == "数学":
            start_col = 19
        elif input_sub == "英语":
            start_col = 27
        elif input_sub == "科学":
            start_col = 35
        else:
            g.msgbox("输入有误，请重试")
            continue
        
        data = []
        for i in range(nrows):
            if df.iloc[i][0] == '' or df.iloc[i][0] == '/':
                continue
            get_str = str(df.iloc[i][0])
            if get_str == input_class:
                for j in range(5):
                    rate_str = str(df.iloc[i,j+start_col])
                    rate_len = len(rate_str)
                    data.append(float(rate_str[0:rate_len-1]))

        if len(data) == 0:
            g.msgbox("输入有误，请重试")
            continue
        mpl.rcParams["font.sans-serif"] = ["SimHei"]
        mpl.rcParams["axes.unicode_minus"] = False

        x = [i+1 for i in range(5)]   
        x_label = ["A等率", "B等率", "C等率", "D等率", "E等率"]    
        plt.bar(x, data, align="center", color=color_sch[color_switch][1], tick_label=x_label)

        plt.xlabel("等第")
        plt.ylabel("比例")

        plt.legend()

        plt.show()
    elif input_type == choices[2]:
        input_grade = int(g.buttonbox(msg="选择年级",title="成绩整理与分析系统",choices=("1","2","3","4","5","6")))
        input_sub = g.buttonbox(msg="选择学科",title="成绩整理与分析系统",choices=("语文","数学","英语","科学"))
        start_col = 0
        if input_sub == "语文":
            start_col = 6
        elif input_sub == "数学":
            start_col = 16
        elif input_sub == "英语":
            start_col = 24
        elif input_sub == "科学":
            start_col = 32
        else:
            g.msgbox("输入有误，请重试")
            continue
        point_tar = 0
        point_dif = []
        class_lst = []
        for i in range(nrows):
            if df.iloc[i][0] == '' or df.iloc[i][0] == '/':
                continue
            first_str = str(df.iloc[i][0])[0]
            if first_str >= '0' and first_str <= '9': #处理班级或年级
                grade = int(first_str)
                if grade == input_grade:
                    if str(df.iloc[i][0])[1] < '0' or str(df.iloc[i][0])[1] > '9': #年级
                        point_tar = float(df.iloc[i][start_col])
                    else: #班级
                        class_lst.append(str(df.iloc[i][0]))
                        point_dif.append(float(df.iloc[i][start_col]))
                elif grade > input_grade:
                    break

        class_cnt = len(class_lst)
        for i in range(class_cnt):
            point_dif[i] = point_dif[i] - point_tar
        mpl.rcParams["font.sans-serif"] = ["SimHei"]
        mpl.rcParams["axes.unicode_minus"] = False

        x = [i+1 for i in range(class_cnt)]   
        plt.bar(x, point_dif, align="center", color=color_sch[color_switch][1], tick_label=class_lst)
        for i in range(class_cnt):
            plt.text(x[i], point_dif[i], "%s" % str(point_dif[i]))

        plt.xlabel("班级")
        plt.ylabel("分差")

        plt.legend()

        plt.show()
    elif input_type == choices[3]:
        input_option = g.buttonbox(msg="更改配置选项",title="成绩整理与分析系统",choices=("修改配色方案","更多功能开发中…"))
        if input_option == "修改配色方案":
            input_color = g.buttonbox(msg="配色方案",title="成绩整理与分析系统",choices=("炫彩色","淡雅暖色", "淡雅冷色"))
            if input_color == "炫彩色":
                color_switch = 0
            elif input_color == "淡雅暖色":
                color_switch = 1
            elif input_color == "淡雅冷色":
                color_switch = 2
            else:
                continue
        else:
            g.msgbox("输入有误，请重试")
            continue
    elif input_type == choices[4]:
        sys.exit(0)
    else:
        g.msgbox("输入有误，请重试")
        continue
