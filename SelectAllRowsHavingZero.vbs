Sub SelectNonBoldRowsWithZero()

    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim i As Long
    startColumn = "C"
    startRow = 8
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("FD4110") ' Change to your sheet name

    ' Define the last row
    lastRow = ws.Cells(ws.Rows.Count, startColumn).End(xlUp).Row

    
    ' Loop through each row, starting from row 8
    For i = startRow To lastRow
        ' Check if the cell in column C is not bold and has a value of zero
        If ws.Cells(i, startColumn).Value = 0 Then
            ' If the conditions are met, select the row
            If rng Is Nothing Then
                Set rng = ws.Rows(i)
            Else
                Set rng = Union(rng, ws.Rows(i))
            End If
        End If
    Next i

    ' Select the rows
    If Not rng Is Nothing Then rng.Select

End Sub

