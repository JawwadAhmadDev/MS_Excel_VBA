Public Function VBA_XLOOKUP(lookup_value As Range, lookup_array As Range, return_array As Range, if_not_found As Variant, Optional match_mode As Long = 0) As Variant
    Dim i As Long
    Dim found As Boolean
    VBA_XLOOKUP = if_not_found ' Set default return value
    For i = 1 To lookup_array.Cells.Count
        Select Case match_mode
            Case 0 ' Exact match
                If lookup_array.Cells(i).Value = lookup_value.Value Then
                    found = True
                End If
            Case -1 ' Exact match or next smaller item
                If lookup_array.Cells(i).Value <= lookup_value.Value Then
                    found = True
                End If
            Case 1 ' Exact match or next larger item
                If lookup_array.Cells(i).Value >= lookup_value.Value Then
                    found = True
                End If
            Case Else ' Unsupported match_mode
                Exit Function
        End Select
        If found Then
            VBA_XLOOKUP = return_array.Cells(i).Value
            Exit Function
        End If
    Next i
End Function
