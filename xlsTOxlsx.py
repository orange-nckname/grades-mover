import win32com.client as win32
import os

def reverse_file(fileName):
    global excel, wb
    try:
        fname = fileName
        fname = os.path.normpath(fname)
        excel = win32.Dispatch('Excel.Application')
        # 是否打开excel界面
        excel.Visible = False
        # 是否关闭保存弹出框
        excel.DisplayAlerts = True
        # 打开当前工作簿
        wb = excel.Workbooks.Open(fname)
        # 保存工作簿
        wb.SaveAs(fname + "x", FileFormat=51)  # FileFormat = 51为xlsx,FileFormat = 56为xls
        # 关闭工作簿
        wb.Close()  #

    except:
        print("error")
    finally:
        # 退出excel,不退出则会被占用
        # wb.Close()  #
        excel.Application.Quit()
        # os.rename(fname, fname + "x")
        return os.path.normpath(fname + "x")

if __name__ == "__main__":
    reverse_file("E:/Desktop/a.xls")
