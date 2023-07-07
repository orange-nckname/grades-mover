"""
Task:
[v] xls转换为xlsx
[v] 输入列数字母
[ ] 添加选择空表格的按钮
[ ] 优化界面文字
[ ] 异常处理
"""


from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import xlsTOxlsx
import os
import sys
from openpyxl import *

root = Tk()

root.title("title")


# 

# def change_fileType():
# TODO
def chose_file():
    global stra
    file_name = filedialog.askopenfilename(filetypes=[("XLS", ".xls")])
    xlsTOxlsx.reverse_file(file_name)
    stra = file_name + "x"
    # print(file_name)

def chose_folder():
    global folder_place
    file_dir = filedialog.askdirectory()
    os.chdir(file_dir)
    folder_place = file_dir
    # print(file_dir)

Label(root, text="选择数据来源文件夹：").grid(row=0)
Button(root, text="选择文件夹", command=chose_folder).grid(row=0, column=2)
Label(root, text="选择带有成绩的表格：").grid(row=2)
Button(root, text="选择文件", command=chose_file).grid(row=2, column=2)

Label(root, text="输入带有成绩的表格里的的学号列\n（大写字母）：").grid(row=4, column=0)

Label(root, text="输入数据来源表格的分数列\n（大写字母）：").grid(row=6, column=0)

student_col_char = StringVar()
student_col = IntVar()
score_col_char = StringVar()
score_col = IntVar()

entry1 = Entry(root, textvariable=student_col_char)
entry2 = Entry(root,  textvariable=score_col_char)


entry1.grid(row=4, column=2, padx=10, pady=5)
entry2.grid(row=6, column=2, padx=10, pady=5)


def show():
    # print(f"学号：{student_col.get()}")
    # print(f"分数：{score_col.get()}")
    root.quit()


Button(root, text="确定", width=10, command=show)\
    .grid(row=8, column=0, sticky=W, padx=10, pady=5)

Button(root, text="取消", width=10, command=sys.exit)\
    .grid(row=8, column=2, sticky=E, padx=10, pady=5)
       



root.mainloop()

# ------------------------------------------------------------------------ #

student_col.set( ord(student_col_char.get()) - 65 )
score_col.set( ord(score_col_char.get()) - 65 )

# print("注：数据储存到fin.xlsx\n数据从00.xlsx读取")
# stra = input("调用表名（不加后缀）：")
# fro = int(input("学号列："))
fro = student_col.get()
# to = int(input("得分："))
to = score_col.get()
# numb = float(input())

res = load_workbook(r"00.xlsx", data_only=True)
a = load_workbook(stra, data_only=True)


table_res = res["Sheet1"]
table_a = a["Sheet1"]


dic_a = {}
dic_b = {}
dic_c = {}
dic_d = {}
dic_e = {}
list_a = []


for i in range(1, table_a.max_row):
    #for item in list(table_a.rows)[i]:  
    dic_a[str(list(table_a.rows)[i][fro].value)] = str(list(table_a.rows)[i][to].value)
    # print("aaa")

for i in range(1, table_res.max_row):
    list_a.append(str(list(table_res.columns)[0][i].value))
    # print("BBB")

# print(dic_a)
# print()
# print()
# print()
# print()
# print(list_a)

keys = list(dic_a.keys())
# print(list_a)

# keys: dic_a:  a表
# list_a:       总表

for i in keys:
    if (i != '' and i != "None" and dic_a[i] != 'None' and dic_a[i] != '-' and dic_a[i] != ''):
        try:
            ind = list_a.index(i)
        except:
            continue
        # print(ind)
            
        table_res.cell(ind+2, 4).value = ""
        table_res.cell(ind+2, 4).value = round(float(dic_a[i]), 0)
        # / numb


res.save(f"{folder_place}\\fin.xlsx")
# ------------------------------------------------------------------------ #

# except Exception as reason:

#     messagebox.showerror(title="运行时错误！", message=f"""程序抛出了一个小小的错误~
# {str(reason)}""")