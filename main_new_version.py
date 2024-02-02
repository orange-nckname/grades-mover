from tkinter import *
from tkinter import messagebox
import sys

# ------------------------------------------------------------------------ #

try:
    from main_model import *
    
# ------------------------------------------------------------------------ #

except TypeError as reason:
    if "ord()" in str(reason):
        sys.exit()
    messagebox.showerror(title="程序运行时出现错误！", message=f"""程序运行时出现错误！
                         
您可能没有选择文件，或者文件打不开。
                
                         
具体错误原因：{reason}""")
except KeyError as reason:
    if str(reason) == "'Worksheet Sheet1 does not exist.'":
        messagebox.showerror(title="程序运行时出现错误！", message=f"""程序运行时出现错误！
                             
您选择的 Excel 文件中的工作标名称可能不是“Sheet1”。
                             
请您在您选择的 Excel 表格中所需要的工作表的名字改为“Sheet1”后再次尝试。
                             

具体错误原因：{reason}""")
except IndexError as reason:
    messagebox.showerror(title="程序运行时出现错误！", message=f"""程序运行时出现错误！
                         
您刚刚在程序界面输入的数据可能有问题。
                         
您只需要在输入框中输入表格对应的编号：一个大写字母即可，无需任何数字、符号、空格。
                         
请检查输入内容后重试。
                         
                         
具体错误原因：{reason}""")
except Exception as reason:
    
    messagebox.showerror(title="程序运行时出现错误！", message=f"""程序运行时出现错误！

您遇到了一个程序没有预料到的错误。

可能由以下原因导致：
                         
1. 文件损坏；
                         
2. 代码出bug了等。


错误原因：{reason}
""")

