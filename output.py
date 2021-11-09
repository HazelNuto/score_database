# -*- coding: UTF-8 -*-
import pandas as pd
import numpy as np
import math
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import xlwt
from pandas import DataFrame
import math
import random
import os
import json
import sys

def output_excel():
    text = open('./data/result.txt',"r").read() 
    line = text.split('\n')

    f = xlwt.Workbook(encoding = 'utf-8') #创建工作薄
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet
    for i in range(len(line)):
        block = line[i].split(' ')
        for j in range(len(block)):
            ss = block[j]
            sheet1.write(i,j,ss)
    outputfile = './data/result.xls'
    f.save(outputfile)#保存文件

output_excel()