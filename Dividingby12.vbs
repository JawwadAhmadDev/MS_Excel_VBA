Sub ConvertNumberToFormula()

    Dim cell As Range
    Dim targetRange As Range
    
    ' Prompt the user to select a range
    On Error Resume Next
    Set targetRange = Application.InputBox("Select a range", Type:=8)
    On Error GoTo 0
    
    ' Exit if no range is selected
    If targetRange Is Nothing Then Exit Sub
    
    For Each cell In targetRange
        ' Check if the cell contains a numeric value
        If IsNumeric(cell.Value) And cell.Value <> 0 Then
            ' Convert to the desired formula (divide by 12 and remove "=")
            cell.Formula = cell.Value / 12
        End If
    Next cell

End Sub
