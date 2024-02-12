Sub SplitTextInColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim textArray() As String
    Dim newText As String
    Dim originalValues As Variant
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with your actual sheet name
    
    ' Find the last row with data in column 8
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    
    ' Store the original values in an array
    originalValues = ws.Range("H2:H" & lastRow).Value
    
    ' Loop through each cell value in column 8
    For i = 1 To UBound(originalValues, 1)
        ' Check if the cell contains the pattern
        If InStr(1, originalValues(i, 1), "Issue:") > 0 And InStr(1, originalValues(i, 1), "Resolution:") > 0 Then
            ' Split the text based on "Resolution:" and create multiple lines
            textArray = Split(originalValues(i, 1), "Resolution:")
            
            ' Format the new text
            newText = "Issue:" & Trim(textArray(0)) & vbCrLf & "Resolution:" & Trim(textArray(1))
            
            ' Update the array with the new text
            originalValues(i, 1) = newText
        End If
    Next i
    
    ' Update the entire column with the reformatted values
    ws.Range("H2:H" & lastRow).Value = originalValues
End Sub
