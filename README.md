# 程序使用说明

## 程序功能

该程序可以将一个 Excel 文档中的学生成绩**按学号内容匹配转存至另一表格**，允许多出 or 缺少某个同学。

## 使用方法

1. 将空白成绩表重命名为 00.xlsx，并复制该表，命名为 fin.xlsx。

> 注：需保证 00.xlsx 和 fin.xlsx 的学号列为 A 列，成绩列为 D 列。

```
文件夹
|___ 00.xlsx
|___ fin.xlsx
|___ ...
```

选择原始成绩表 --> 00.xlsx（空白成绩表） --> fin.xlsx（最终成绩表）



2. 先运行该程序，弹出以下窗口：

![](C:\Users\cch12\AppData\Roaming\marktext\images\2022-08-11-18-45-09-image.png)

> 选择数据来源文件夹：就是 0.xlsx 和 fin.xlsx 所在的文件夹；
> 
> 选择数据来源表格：指原始成绩表；
> 
> 输入数据来源表格的学号列（从0开始）：指**原始数据**的**学号**的列数，第一列为 0；
> 
> 输入数据来源表格的分数列（从0开始）：指**原始数据**的**成绩**的列数，第一列为 0。



3. 点击确定按钮，稍等片刻窗口会自动关闭。可在 fin.xlsx 中查看结果。


