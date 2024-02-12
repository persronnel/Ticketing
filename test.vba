Sub ReformatData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with your actual sheet name
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the rows and reformat the data
    For i = 2 To lastRow ' Assuming data starts from the second row
        ' Priority
        ws.Cells(i, 1).Value = Left(ws.Cells(i, 1).Value, 2)
        
        ' End User
        ws.Cells(i, 3).Value = Mid(ws.Cells(i, 3).Value, InStr(1, ws.Cells(i, 3).Value, " ") + 1)
        
        ' State
        ws.Cells(i, 4).Value = ws.Cells(i, 5).Value
        
        ' Summary
        ws.Cells(i, 5).Value = ws.Cells(i, 6).Value
        
        ' Opened
        ws.Cells(i, 6).Value = Format(ws.Cells(i, 7).Value, "dd-mmm-yyyy")
        
        ' Resolved
        ws.Cells(i, 7).Value = Format(ws.Cells(i, 8).Value, "dd-mmm-yyyy")
        
        ' Comments
        ws.Cells(i, 8).Value = "[" & Mid(ws.Cells(i, 9).Value, InStr(1, ws.Cells(i, 9).Value, "]") + 1)
    Next i
End Sub
