# Grades-Mover App
[TOC]
## 功能：
选择两个Excel文件，将一个中的数据内容根据指定列对应到另外一个文件中。  
支持.xls、.xlsx。

## 使用方法：
程序组成：  
+ 源代码：
```
grades-mover
|---main.py         # 主程序
|---main_model.py   # 处理主程序
|---xlsTOxlsx.py    # 将.xls转换为.xlsx
```
+ 打包后程序：
```
grades-mover
|---main.exe        # 主程序
|---excel.ico       # 图标文件
```
依赖库文件：
+ openpyxl

使用方法：
双击运行main.exe文件。
![content](https://api.onedrive.com/v1.0/shares/s!Ar8BtB6LV-uigV9P88A5Vl8sCWA9/root/content)

根据界面选择两个文件、填入相应内容后点击确定。  
在当前目录下生成output.xlsx和status.txt，分别对应生成的excel文件和生成过程中遇到的问题（例如无法录入、不存在等）并自动打开。