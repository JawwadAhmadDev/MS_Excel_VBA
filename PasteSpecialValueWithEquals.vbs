Sub PasteSpecialValueWithEquals()

    Dim rng As Range
    Dim cell As Range

    ' Ensure that there's a valid selection to paste into
    If Not Selection Is Nothing Then
        On Error Resume Next ' Resume on next line in case of error

        ' Paste the results (values) of formulas
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        On Error GoTo 0 ' Reset error handling to default behavior

        ' Loop through each cell in the selection and prefix the value with "="
        Set rng = Selection
        For Each cell In rng
            If cell.Value <> "" Then
                cell.Value = "=" & cell.Value
            End If
        Next cell
    End If

End Sub

