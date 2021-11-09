# -*- coding: UTF-8 -*-
import pandas as pd
import numpy as np
import math
import matplotlib.pyplot as plt
import numpy as np
import xlrd
from pandas import DataFrame
import xlwt
import math
import random
import os
import json
import sys

dataPath = "./data/data.xlsx"
o = dataPath
wb = xlrd.open_workbook(dataPath)
sheets = wb.sheet_names()
outputfile = "./data/result.txt"

class Logger(object):
    def __init__(self, filename="Default.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "a")
 
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
 
    def flush(self):
        pass

def cal_bef_end(score, cnt):
    if cnt == 0:
        return [0,0]
    tp_sum = 0
    score.sort(reverse = True)
    before = math.floor(cnt*0.45)
    end = math.floor(cnt*0.3)
    for i in range(before):
        tp_sum += score[i]
    res_bef = tp_sum / before

    score.sort()
    tp_sum = 0.0
    for i in range(end):
        tp_sum += score[i]
    res_end = tp_sum / end
    return [res_bef, res_end]

def cal_grade_levelscore(score, cnt):
    if cnt == 0:
        return [0,0,0,0]
    a_num = math.floor(cnt * 0.2)
    b_num = math.floor(cnt * 0.45)
    c_num = math.floor(cnt * 0.7)
    d_num = math.floor(cnt * 0.9)
    return [score[a_num], score[b_num], score[c_num], score[d_num]]

def cal_level(score, level):
    rows = len(score)
    if rows == 0:
        return [0,0,0,0,0]
    cnt = [0, 0, 0, 0, 0]
    score.sort(reverse = True)
    for point in score:
        if point >= level[0]:
            cnt[0] += 1
        elif point < level[0] and point >= level[1]:
            cnt[1] += 1
        elif point < level[1] and point >= level[2]:
            cnt[2] += 1
        elif point < level[2] and point >= level[3]:
            cnt[3] += 1
        else:
            cnt[4] += 1
    return [cnt[i]*100.0/rows for i in range(5)]

def print_title():
    print("班级 原有人数 现有人数 语文 / / / / / / / / / / 数学 / / / / / / / / / 英语 / / / / / / / 科学 / / / / / / /")
    print("/ / / 基础 阅读 写作 总分 前45% 后30% A等率 B等率 C等率 D等率 E等率 知识技能 综合应用 总分 前45% 后30% A等率 B等率 C等率 D等率 E等率 总分 前45% 后30% A等率 B等率 C等率 D等率 E等率 总分 前45% 后30% A等率 B等率 C等率 D等率 E等率")

df_28 = DataFrame()
path = os.path.abspath(os.path.dirname(__file__))
type = sys.getfilesystemencoding()
sys.stdout = Logger(outputfile)

inputfile = pd.read_excel(dataPath, sheet_name=None)
class_lst = list(inputfile)
class_cnt = len(class_lst)-1 #班级数
class_rows = [0 for i in range(class_cnt)]
class_c = [[0,0,0,0] for i in range(class_cnt)]
class_m = [[0,0,0]   for i in range(class_cnt)]
class_e = [0 for i in range(class_cnt)]
class_s = [0 for i in range(class_cnt)]
class_c_cnt = [0 for i in range(class_cnt)]
class_m_cnt = [0 for i in range(class_cnt)]
class_e_cnt = [0 for i in range(class_cnt)]
class_s_cnt = [0 for i in range(class_cnt)]
class_c_score = [[] for i in range(class_cnt)] #各个班级的所有语文总分
class_m_score = [[] for i in range(class_cnt)] #各个班级的所有数学总分
class_e_score = [[] for i in range(class_cnt)] #各个班级的所有英语总分
class_s_score = [[] for i in range(class_cnt)] #各个班级的所有科学总分

grade_c = [0 for i in range(12)] # 第i<<1位代表总分，i<<1|1位代表参考人数
grade_m = [0 for i in range(12)]
grade_e = [0 for i in range(12)]
grade_s = [0 for i in range(12)]
grade_c_score = [[] for i in range(6)] #六个年级的所有语文总分
grade_m_score = [[] for i in range(6)] #六个年级的所有数学总分
grade_e_score = [[] for i in range(6)] #六个年级的所有英语总分
grade_s_score = [[] for i in range(6)] #六个年级的所有科学总分
grade_c_level = [[0 for i in range(4)] for i in range(6)]
grade_m_level = [[0 for i in range(4)] for i in range(6)]
grade_e_level = [[0 for i in range(4)] for i in range(6)]
grade_s_level = [[0 for i in range(4)] for i in range(6)]

for i in range(0, class_cnt): #处理和存储
    df = pd.read_excel(dataPath, sheet_name=i, skiprows=0,keep_default_na=False) #依次打开所有sheet
    grade = int(str(class_lst[i])[0]) - 1

    nrows = 0 #班级人数
    c = [0, 0, 0, 0] #语文基础知识、阅读、作文、总分
    m = [0, 0, 0]    #数学知识技能、综合应用、总分
    e = 0  #英语总分
    s = 0  #科学总分
    pres = [0, 0, 0, 0] #语文、数学、英语、科学参考人数

    for j in range(1, df.shape[0]):
        if df.iloc[j, 0] == '': 
            break
        nrows += 1 #班级人数
        if df.iloc[j, 2] != '': #语文
            pres[0] += 1
            for dx in range(2,6):
                c[dx-2] += float(df.iloc[j,dx])
            class_c_score[i].append(float(df.iloc[j,5]))
            grade_c_score[grade].append(float(df.iloc[j,5]))
        if df.iloc[j, 6] != '': #数学
            pres[1] += 1
            for dx in range(6,9):
                m[dx-6] += float(df.iloc[j,dx])
            class_m_score[i].append(float(df.iloc[j,8]))
            grade_m_score[grade].append(float(df.iloc[j,8]))
        if df.shape[1] > 9:         #要考英科
            if df.iloc[j, 9] != '': #英语
                pres[2] += 1
                e += float(df.iloc[j,9])
                class_e_score[i].append(float(df.iloc[j,9]))
                grade_e_score[grade].append(float(df.iloc[j,9]))
            if df.iloc[j,10] != '': #科学
                pres[3] += 1
                s += float(df.iloc[j,10])
                class_s_score[i].append(float(df.iloc[j,10]))
                grade_s_score[grade].append(float(df.iloc[j,10]))
    class_rows[i] = nrows
    for j in range(4):
        class_c[i][j] = c[j]
    for j in range(3):
        class_m[i][j] = m[j]
    class_e[i] = e
    class_s[i] = s
    class_c_cnt[i] = pres[0]
    class_m_cnt[i] = pres[1]
    class_e_cnt[i] = pres[2]
    class_s_cnt[i] = pres[3]
    #赋值班级

    grade_c[grade<<1] += c[3]
    grade_c[grade<<1|1] += pres[0]
    grade_m[grade<<1] += m[2]
    grade_m[grade<<1|1] += pres[1]
    grade_e[grade<<1] += e
    grade_e[grade<<1|1] += pres[2]
    grade_s[grade<<1] += s
    grade_s[grade<<1|1] += pres[3]
    #赋值年级

for i in range(6): #划分等第
    grade_c_score[i].sort(reverse = True)
    grade_m_score[i].sort(reverse = True)
    grade_e_score[i].sort(reverse = True)
    grade_s_score[i].sort(reverse = True)
    grade_c_level[i] = cal_grade_levelscore(grade_c_score[i], grade_c[i<<1|1])
    grade_m_level[i] = cal_grade_levelscore(grade_m_score[i], grade_m[i<<1|1])
    grade_e_level[i] = cal_grade_levelscore(grade_e_score[i], grade_e[i<<1|1])
    grade_s_level[i] = cal_grade_levelscore(grade_s_score[i], grade_s[i<<1|1])

def output_res():
    grade = 1
    print_title()
    for i in range(0, class_cnt): #计算和输出
        new_grade = int(str(class_lst[i])[0]) - 1
        if grade != new_grade: #一个年级处理完
            print("%d年级" % (grade+1), end = ' ')
            print("/ / / / / %.2f / / / / / / / / / %.2f" % ( grade_c[grade<<1] / grade_c[grade<<1|1], grade_m[grade<<1] / grade_m[grade<<1|1] ), end = ' ')
            if grade_e[grade<<1|1] != 0:
                print("/ / / / / / / %.2f" % (grade_e[grade<<1] / grade_e[grade<<1|1]), end = ' ')
            if grade_s[grade<<1|1] != 0:
                print("/ / / / / / / %.2f" % (grade_s[grade<<1] / grade_s[grade<<1|1]), end = ' ')
            print("\n")
            print_title()
        grade = new_grade

        c_bef_end = cal_bef_end(class_c_score[i], class_c_cnt[i])
        m_bef_end = cal_bef_end(class_m_score[i], class_m_cnt[i])
        e_bef_end = cal_bef_end(class_e_score[i], class_e_cnt[i])
        s_bef_end = cal_bef_end(class_s_score[i], class_s_cnt[i])
        c_level = cal_level(class_c_score[i], grade_c_level[grade])
        m_level = cal_level(class_m_score[i], grade_m_level[grade])
        e_level = cal_level(class_e_score[i], grade_e_level[grade])
        s_level = cal_level(class_s_score[i], grade_s_level[grade])

        print("%s %d %d" % ( class_lst[i], class_rows[i], max(max(class_c_cnt[i],class_m_cnt[i]),max(class_e_cnt[i],class_s_cnt[i])) ), end=' ') #班级编号，班级人数，实际人数
        print("%.2f %.2f %.2f %.2f %.2f %.2f" % ( class_c[i][0] / class_c_cnt[i], class_c[i][1] / class_c_cnt[i], class_c[i][2] / class_c_cnt[i], class_c[i][3] / class_c_cnt[i], c_bef_end[0], c_bef_end[1] ), end =' ') #语文分值
        for j in range(5):
            print("%.2f" % (c_level[j]), end = '% ')
        print("%.2f %.2f %.2f %.2f %.2f" % ( class_m[i][0] / class_m_cnt[i], class_m[i][1] / class_m_cnt[i], class_m[i][2] / class_m_cnt[i], m_bef_end[0], m_bef_end[1] ), end = ' ') #数学分值
        for j in range(5):
            print("%.2f" % (m_level[j]), end = '% ')
        if class_e_cnt[i] != 0:
            print("%.2f %.2f %.2f" % ( class_e[i] / class_e_cnt[i], e_bef_end[0], e_bef_end[1] ), end = ' ')
            for j in range(5):
                print("%.2f" % (e_level[j]), end = '% ')
        if class_s_cnt[i] != 0:
            print("%.2f %.2f %.2f" % ( class_s[i] / class_s_cnt[i], s_bef_end[0], s_bef_end[1] ), end = ' ')
            for j in range(5):
                print("%.2f" % (s_level[j]), end = '% ')
        print()

    print("%d年级" % (grade+1), end = ' ')
    print("/ / / / / %.2f / / / / / / / / / %.2f" % ( grade_c[grade<<1] / grade_c[grade<<1|1], grade_m[grade<<1] / grade_m[grade<<1|1] ), end = ' ')
    if grade_e[grade<<1|1] != 0:
        print("/ / / / / / / %.2f" % (grade_e[grade<<1] / grade_e[grade<<1|1]), end = ' ')
    if grade_s[grade<<1|1] != 0:
        print("/ / / / / / / %.2f" % (grade_s[grade<<1] / grade_s[grade<<1|1]), end = ' ')
    print("\n")

output_res()