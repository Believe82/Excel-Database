'Code to open the external file:

Sub open_wb_onedrive()
'Opens an instance of our data workbook

Dim wbFullName

Set objLogExcel = CreateObject("Excel.Application")
objLogExcel.Visible = True

'Put here the path to the workbook you want to be opened
wbFullName = "C:\Users\data.xlsx"
Set openwb = objLogExcel.Workbooks.Open(wbFullName)
   
End Sub
