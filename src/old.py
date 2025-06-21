import win32com.client

BookName = "sample.xlsm"

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(r"C:\Users\IKens\Desktop\05_python\dev\xlsm\sample.xlsm")
excel.Application.Run("Test")
wb.Close(SaveChanges=True)
excel.Quit()


