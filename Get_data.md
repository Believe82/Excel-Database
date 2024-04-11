Code to get data:

Sub get_data()
Dim PasteWs As Worksheet
Dim CopyWs As Worksheet
Dim LastRow
Dim LastColumn
'Code that will get the data:

'Opens the other workbook and sets the sheet to copy from
open_wb_onedrive
Set CopyWs = openwb.Worksheets("data")

'Set the sheet you would like to paste into:
Set PasteWs = ThisWorkbook.Worksheets("Main")

'clear contents of the paste area to make copy easier
PasteWs.Range("A1:AA100000").ClearContents

Application.ScreenUpdating = False

With CopyWs
    
    'Uncomment the following if you are using Passwords:
    '.Unprotect Password:=pass
    
    'Finds the last Row and last Coloumn of our data
    LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    
    'copy and paste the data from our data workbook to our main one
    .Range(.Cells(1, 1), .Cells(LastRow, LastColumn)).Copy
    PasteWs.Activate
    PasteWs.Range("A1").Select
    PasteWs.Paste
    
    'Uncomment the following if you are using Passwords:
    '.Protect Password:=pass
    
End With

'closes the data workbook and saves
openwb.Save
openwb.Close
objLogExcel.Quit

Application.ScreenUpdating = True

End Sub

