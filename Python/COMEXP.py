##A COM server is a software (DLL/EXE) that accepts remote procedure calls (RPC)

from win32com import client

Excel = client("Excel.Application")
Excel.Visible = True
wb = Excel.Workbooks.add()
ws = wb.Worksheets.add()
ws.Name = "My Worksheet"
ws.Range("A1:A1").Value = "Hello World"
ws.SaveAs("C:\xlsx.xlsx)
Excel.Application.Quit()