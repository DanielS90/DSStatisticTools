Option Compare Text

Public Function DS_CountIfAny(ByVal cellRange As Range, ByVal comp1 As Variant, Optional comp2 As Variant, Optional comp3 As Variant, Optional comp4 As Variant)
    Dim currentCell As Range
    
    If Not IsMissing(comp4) Then
        For Each currentCell In cellRange
            If currentCell Like comp1 Or currentCell Like comp2 Or currentCell Like comp3 Or currentCell Like comp4 Then DS_CountIfAny = DS_CountIfAny + 1
        Next currentCell
    ElseIf Not IsMissing(comp3) Then
        For Each currentCell In cellRange
            If currentCell Like comp1 Or currentCell Like comp2 Or currentCell Like comp3 Then DS_CountIfAny = DS_CountIfAny + 1
        Next currentCell
    ElseIf Not IsMissing(comp2) Then
        For Each currentCell In cellRange
            If currentCell Like comp1 Or currentCell Like comp2 Then DS_CountIfAny = DS_CountIfAny + 1
        Next currentCell
    Else
        For Each currentCell In cellRange
            If currentCell Like comp1 Then DS_CountIfAny = DS_CountIfAny + 1
        Next currentCell
    End If
End Function

Public Function DS_CommaListAsArray(ByVal cellRange As Range)
    Dim valueArray() As Variant
    ReDim valueArray(0)
    Dim counter As Integer
    counter = 0
    For Each currentCell In cellRange
        Dim values() As String
        values = split(currentCell, ",")
        For Each value In values
            ReDim Preserve valueArray(0 To counter)
            valueArray(counter) = value
            counter = counter + 1
        Next value
    Next currentCell
    
    DS_CommaListAsArray = valueArray
End Function

Public Function DS_CommaListCountValue(ByVal cellRange As Range, ByVal comp As Variant)
    Dim valueArray() As Variant
    valueArray = DS_CommaListAsArray(cellRange)
    
    Dim val As Variant
    For Each val In valueArray
        If val Like comp Then DS_CommaListCountValue = DS_CommaListCountValue + 1
    Next val
End Function

Public Function DS_CommaListValueAt(ByVal cellRange As Range, ByVal index As Integer, Optional decimalSeparator As Variant)
    If IsMissing(decimalSeparator) Then
        decimalSeparator = "."
    End If
    
    Dim value As Variant
    value = DS_CommaListAsArray(cellRange)(index)
    
    If DS_IsNumeric(Replace(value, decimalSeparator, Application.decimalSeparator)) Then
        value = DS_StringToNumber(value, decimalSeparator)
    End If
    
    DS_CommaListValueAt = value
End Function

Public Function DS_RangeToArray(ByVal cellRange As Range)
    Dim valueArray() As Variant
    ReDim valueArray(0)
    Dim counter As Integer
    counter = 0
    For Each currentCell In cellRange
        ReDim Preserve valueArray(0 To counter)
        valueArray(counter) = currentCell.value
        counter = counter + 1
    Next currentCell
    DS_RangeToArray = valueArray
End Function

Public Function DS_CommaListMake(ByVal cellRange As Range, Optional decimalSeparator As Variant)
    Dim valueArray() As Variant
    valueArray = DS_RangeToArray(cellRange)
    
    If IsMissing(decimalSeparator) Then
        decimalSeparator = "."
    End If
    
    Dim result As String
    For Each value In valueArray
        If Len(result) > 0 Then
            result = result & ","
        End If
        Dim strVal As String
        strVal = CStr(value)
        If DS_IsNumeric(value) Then
            strVal = Replace(strVal, Application.thousandsSeparator, "")
            strVal = Replace(strVal, Application.decimalSeparator, decimalSeparator)
        End If
        
        result = result & strVal
    Next value
    
    DS_CommaListMake = result
End Function
