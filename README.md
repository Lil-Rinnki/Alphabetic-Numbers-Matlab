# Matlab 数字转换为Excel字母列号

Convert numeric column numbers to Excel alphabetic column numbers

## 一、前言

在对海量数据进行处理时，往往导入的txt或xls格式文件会用每一列表达并储存不同的变量。

观察可知Excel中为直观区分行号与列号，行号用从1开始的数字表示，列号用A开始的字母表示。特别地，当记录物理量多至一定程度（26列），列号以二十六进制进行增长，如第27列的列号为AA。

在Matlab中读取数据文件常使用的函数为readtable，可以使用一下代码对特定表格文件中列号为column\_number的列进行提取。

```matlab
T = readtable(filename,'Range','column_number：column_number')
```

实际提取过程中，往往得到的是该列号的数字序号，出于此目的开发了本数字列号转字母列号函数num2letter。

## 二、代码

```matlab
function col_letter = num2letter(col_num)

string = {'A','B','C','D','E','F',...
          'G','H','I','J','K','L',...
          'M','N','O','P','Q','R',...
          'S','T','U','V','W','X','Y','Z'};

col_letter = [];

while col_num>0

    m = mod(col_num,26);

    if m == 0
        m = 26;
    end

    col_letter = [string{m} col_letter];

    col_num = (col_num-m)/26;
end

end
```

## 三、代码解读

num2letter函数将列号转化为Excel表格中对应字母编号的函数。其输入值为col\_num，是目标列的数字序号，double型；输出值col\_letter为输入值在Excel表格中对应字母列号。其主要思想为二十六进制与取余计算。

如想求第24列在Excel表格对应字母列号，可用如下代码运算：

```matlab
colName1 = ExcelCol(24);

>> colName

>> ans = 'X';
```

可知第24列对应字母X列。

同理，想查看第252231列在Excel表格对应列，可用如下代码运算：

```matlab
colName1 = ExcelCol(252231);

>> colName

>> ans = 'NICE';
```

可知第252231列对应字母第NICE列。

特别鸣谢：CSDN博主「不爱学习的笨蛋」

在该代码基础上进行了简单修改，无需使用fliplr倒序函数

原文链接：[https://blog.csdn.net/u011624019/article/details/89221789](https://blog.csdn.net/u011624019/article/details/89221789 "https://blog.csdn.net/u011624019/article/details/89221789")
