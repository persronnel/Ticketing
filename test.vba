Sub ReformatData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet (change "Sheet1" to your actual sheet name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row from 2 to the last row
    For i = 2 To lastRow
        ' Format the data in the desired way
        ws.Cells(i, 1).Value = Left(ws.Cells(i, 1).Value, 2) ' Extract first two characters for Priority
        ws.Cells(i, 6).Value = Format(ws.Cells(i, 6).Value, "dd-mmm-yyyy") ' Format Opened date
        ws.Cells(i, 7).Value = Format(ws.Cells(i, 8).Value, "dd-mmm-yyyy") ' Format Resolved date
        ws.Cells(i, 8).Value = "[" & Mid(ws.Cells(i, 7).Value, 14, 20) & "]" & Mid(ws.Cells(i, 9).Value, 12) ' Format Comments
        
        ' Rearrange columns and delete unnecessary columns
        ws.Rows(i).Columns("B:D").Cut
        ws.Rows(i).Columns("A").Insert Shift:=xlToRight
        ws.Rows(i).Columns("H:I").Cut
        ws.Rows(i).Columns("F").Insert Shift:=xlToRight
        ws.Rows(i).Columns("C:D").Delete Shift:=xlToLeft
    Next i
    
    ' Autofit columns for better visibility
    ws.Rows(1).EntireColumn.AutoFit
End Sub
