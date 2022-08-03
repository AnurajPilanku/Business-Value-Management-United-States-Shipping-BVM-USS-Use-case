#Anuraj Pilanku
from win32com.client import Dispatch
import sys
import openpyxl
wkbk1 = sys.argv[1]

wb=openpyxl.load_workbook(wkbk1)
ws=wb.active
rowcount=ws.max_row
wb.close()

try:
    excel = Dispatch("Excel.Application")
    excel.Visible = 0
    source = excel.Workbooks.Open(wkbk1)
    # sh=source.worksheets(0)
    # used=sh.UsedRange
    excel.Range("P2:P"+str(rowcount)).Select()
    excel.Selection.Copy()
    # copy = excel.Workbooks.Open(wkbk2)
    excel.Range("P2:P"+str(rowcount)).Select()
    excel.Selection.PasteSpecial(Paste=-4163)
    source.Save()  # SaveAs(Filename:=sys.argv[1])
    source.Close()
    excel.Quit()
    print("success")
except:
    import os
    import win32com.shell.shell as shell
    import time

    # commands='taskkill /f /im EXCEL.EXE'
    # shell.ShellExecuteEx(lpVerb='runas',lpFile='cmd.exe',lpParameters='/c'+commands)
    os.system('taskkill /f /im EXCEL.EXE')
    time.sleep(10)

    excel = Dispatch("Excel.Application")
    excel.Visible = 0
    source = excel.Workbooks.Open(wkbk1)
    # sh=source.worksheets(0)
    # used=sh.UsedRange
    excel.Range("P2:P"+str(rowcount)).Select()
    excel.Selection.Copy()
    # copy = excel.Workbooks.Open(wkbk2)
    excel.Range("P2:P"+str(rowcount)).Select()
    excel.Selection.PasteSpecial(Paste=-4163)
    source.Save()  # SaveAs(Filename:=sys.argv[1])
    source.Close()
    excel.Quit()
    print("success")
else:
    print("Excel instances opened in server causing issues")
finally:
    print("Excel instances opened in server causing issues")



#python ship.py  C:\Users\2040664\anuraj\ipi2k\test.xlsx  C:\Users\2040664\anuraj\ipi2k\testyt.xlsx