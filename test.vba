Sub SplitTextInColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim textArray() As String
    Dim newText As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with your actual sheet name
    
    ' Find the last row with data in column 8
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    
    ' Loop through each cell in column 8
    For Each cell In ws.Range("H2:H" & lastRow)
        ' Check if the cell contains the pattern
        If InStr(1, cell.Value, "Issue:") > 0 And InStr(1, cell.Value, "Resolution:") > 0 Then
            ' Split the text based on "Resolution:" and create multiple lines
            textArray = Split(cell.Value, "Resolution:")
            
            ' Format the new text
            newText = "Issue:" & Trim(textArray(0)) & vbCrLf & "Resolution:" & Trim(textArray(1))
            
            ' Update the cell with the new text
            cell.Value = newText
        End If
    Next cell
End Sub
