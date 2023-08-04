Sub SelectRowsWithAllZeros()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim startRow As Long, startColumn As Long, endColumn As Long
    Dim allZeros As Boolean

    startRow = 15
    startColumn = 5  ' Column E
    endColumn = 22  ' Column J, change this to your ending column

    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("LQ4028-(1st)") ' Change to your sheet name

    ' Define the last row
    lastRow = ws.Cells(ws.Rows.Count, startColumn).End(xlUp).Row

    ' Loop through each row
    For i = startRow To lastRow
        ' Start with the assumption that all cells in the row are zero
        allZeros = True
        ' Loop through each column in the range
        For j = startColumn To endColumn
            ' Check if the cell is not zero
            If ws.Cells(i, j).Value <> 0 Then
                ' If any cell is not zero, set allZeros to False and exit the column loop
                allZeros = False
                Exit For
            End If
        Next j
        ' If all cells in the row were zero, select the row
        If allZeros Then
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
