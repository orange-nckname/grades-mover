import win32com.client as win32

def reverse_file(fileName):
    fname = fileName
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
    wb.Close() #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

reverse_file("C:\\Users\\thisi\\Desktop\\康复2班 - 副本.xls")

