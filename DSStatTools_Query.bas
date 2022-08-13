Option Compare Text

Public Function DS_Select(ByVal cellRange As Range, ByVal conditionRange As Range, ByVal comparison As String)
    If Not cellRange.Rows.Count = conditionRange.Rows.Count Then
        Exit Function
    End If
    
    Dim result() As Variant
    
    Dim iR As Integer
    Dim counter As Integer
    counter = 0
    For iR = 1 To cellRange.Rows.Count
        If conditionRange(iR, 1) Like comparison Then
            ReDim Preserve result(0 To counter)
            result(counter) = cellRange(iR, 1)
            counter = counter + 1
        End If
    Next iR
    
    DS_Select = result
End Function
