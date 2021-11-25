### 用于全校成绩整理与分析(v2.0)

----

### 功能介绍

* 自动汇总全校考试成绩，自动计算生成前45%、后30%平均分、ABCDE等率，生成excel数据表
* 对汇总数据进行数据分析，对各年级、各班级考试情况作出考评，对分析生成数据可视化图表

#### 数据输入

* 以./data目录下的'data.xlsx'为模板，用在线收集的方式汇总
* 注意：**成绩必须为数字，有特殊情况（如缺考等）也不要填写0或文字，直接跳过，不要填写任何内容**

#### 环境依赖

* python3.0以上，安装pandas，matplotlib，openpyxl库
* 安装方式：

```
pip install pandas
pip install matplotlib
pip install openpyxl
```



#### 使用方法

* cd进入相应目录，运行main.py

```
python main.py
```

* 单击“数据汇总”按钮：在./data目录下会生成result.data文件，就是汇总后生成的全校数据
* 单击“数据分析”按钮：对汇总后生成的数据进行计算分析和可视化呈现



#### 问题解决

* 错误信息：FileNotFoundError: [Errno 2] No such file or directory: './data/data.xlsx'
  * 原因：找不到待处理的数据源文件
  * 解决办法：检查是否将源文件放入data目录下，并改名为"data.xlsx"；注意：拓展名为xlsx，如果为旧版本的xls，要打开文件后，另存为为xlsx后缀的文件。
* 错误信息：ValueError: could not convert string to float: '缺考'
  * 原因：在本该填入成绩的单元格中填入中文
  * 解决办法：检查所有数据表，确保填入成绩的单元格中不出现中文；如有缺考、请假等特殊情况，不要填入任何内容



----

copyright ©王泽宇 HazelNut | shiraihazel@gmail.com | 1023982133@qq.com

blog: https://github.com/HazelNuto | https://www.cnblogs.com/hazelnut

tel: 13588363606

Thanks for your download and welcome to contact me:)

