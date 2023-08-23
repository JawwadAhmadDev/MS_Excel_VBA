Sub PasteSpecialValueOfXLOOKUPWithEqual()

    Dim rng As Range
    Dim cell As Range

    ' Ensure that there's a valid selection
    If Not Selection Is Nothing Then
        Set rng = Selection
        
        ' Loop through each cell in the selection
        For Each cell In rng
            ' Check if the cell contains an XLOOKUP formula
            If InStr(1, cell.Formula, "XLOOKUP", vbTextCompare) > 0 Then
                cell.Value = "=" & cell.Value
            End If
        Next cell
    End If

End Sub

