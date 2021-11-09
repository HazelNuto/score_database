### 用于全校成绩整理与分析

----

#### 数据输入

* 以./data目录下的'data.xlsx'为模板，用在线收集的方式汇总
* 注意：**成绩必须为数字，有特殊情况（如缺考等）也不要填写0或文字，直接跳过，不要填写任何内容**

#### 环境依赖

* python3.0以上，安装pandas，numpy，matplotlib库
* 安装方式：

```
pip install pandas
pip install numpy
pip install matplotlib
```



#### 使用方法

* cd进入相应目录，运行getres.py

```
python getres.py
```

* 在./data目录下会生成result.txt文件，再继续运行

```
python output.py
```

* 在./data目录生成result.xlsx文件，为统计汇总结果的表格
* 运行paint.py

```
python paint.py
```

* 可以进行相应对象的数据可视化分析



#### 问题解决

* 输出文件未改写：将*result.txt*和*result.xlsx*内容进行清空，或直接删除该文件。
* 输出错误：检查*data.xlsx*中有无格式错误；如有缺考，请勿填写任何内容至相应的单元格



----

copyright ©王泽宇 HazelNut | shiraihazel@gmail.com | 1023982133@qq.com

blog: https://github.com/HazelNuto | https://www.cnblogs.com/hazelnut

tel: 13588363606

Thanks for your download and welcome to contact me:)

