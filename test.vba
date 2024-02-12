Sub ReformatData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet where your data is located
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row starting from the second row (assuming headers are in the first row)
    For i = 2 To lastRow
        ' Extract data from the existing columns
        Dim priority As String
        Dim number As Long
        Dim endUser As String
        Dim state As String
        Dim summary As String
        Dim opened As Date
        Dim resolved As Date
        Dim comments As String
        
        priority = Left(ws.Cells(i, 1).Value, 2) ' Extract the first two characters
        number = ws.Cells(i, 2).Value
        endUser = ws.Cells(i, 3).Value
        state = ws.Cells(i, 4).Value
        summary = Mid(ws.Cells(i, 6).Value, InStr(ws.Cells(i, 6).Value, "(") + 1, InStr(ws.Cells(i, 6).Value, ")") - InStr(ws.Cells(i, 6).Value, "(") - 1)
        opened = DateValue(ws.Cells(i, 7).Value)
        resolved = DateValue(ws.Cells(i, 8).Value)
        comments = Replace(Replace(Mid(ws.Cells(i, 9).Value, InStr(ws.Cells(i, 9).Value, "(") + 1), ")", ""), "=>", "=>")
        
        ' Output the reformatted data to new columns
        ws.Cells(i, 1).Value = priority
        ws.Cells(i, 2).Value = number
        ws.Cells(i, 3).Value = endUser
        ws.Cells(i, 4).Value = state
        ws.Cells(i, 5).Value = summary
        ws.Cells(i, 6).Value = Format(opened, "d-mmm-yyyy")
        ws.Cells(i, 7).Value = comments
        ws.Cells(i, 8).ClearContents ' Clear unnecessary columns (e.g., old "Requested for Name" column)
        ws.Cells(i, 9).ClearContents
        ws.Cells(i, 10).ClearContents
        ws.Cells(i, 11).ClearContents
    Next i
End Sub
