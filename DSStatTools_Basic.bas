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

Public Function DS_CommaListAsArray(ByVal cellRange As Range, Optional castNumber As Variant)
    Dim valueArray() As Variant
    ReDim valueArray(0)
    Dim counter As Integer
    counter = 0
    For Each currentCell In cellRange
        Dim values() As String
        values = split(currentCell, ",")
        For Each value In values
            ReDim Preserve valueArray(0 To counter)
            If IsMissing(castNumber) Then
                valueArray(counter) = value
            Else
                valueArray(counter) = DS_StringToNumber(value, ".")
            End If
            counter = counter + 1
        Next value
    Next currentCell
    
    DS_CommaListAsArray = valueArray
End Function

Public Function DS_CSVStringAsArray(ByVal csvString As String, Optional castNumber As Variant)
    Dim valueArray() As Variant
    ReDim valueArray(0)
    Dim counter As Integer
    counter = 0
    
    Dim values() As String
    values = split(csvString, ",")
    For Each value In values
        ReDim Preserve valueArray(0 To counter)
        If IsMissing(castNumber) Then
            valueArray(counter) = value
        Else
            valueArray(counter) = DS_StringToNumber(value, ".")
        End If
        counter = counter + 1
    Next value
    
    DS_CSVStringAsArray = valueArray
End Function

Public Function DS_CommaListCountValue(ByVal cellRange As Range, ByVal comp As Variant)
    Dim valueArray() As Variant
    valueArray = DS_CommaListAsArray(cellRange)
    
    Dim val As Variant
    For Each val In valueArray
        If val Like comp Then DS_CommaListCountValue = DS_CommaListCountValue + 1
    Next val
End Function

Public Function DS_CommaListCountEntries(ByVal cellRange As Range)
    Dim valueArray() As Variant
    valueArray = DS_CommaListAsArray(cellRange)
    
    DS_CommaListCountEntries = UBound(valueArray) - LBound(valueArray) + 1
    If DS_CommaListCountEntries = 1 And IsEmpty(valueArray(0)) Then
        DS_CommaListCountEntries = 0
    End If
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
        If Not IsEmpty(currentCell) Then
            ReDim Preserve valueArray(0 To counter)
            valueArray(counter) = currentCell.value
            counter = counter + 1
        End If
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

Public Function DS_PrintIQR(ByVal cellRange As Variant, Optional decimals As Variant)
    If TypeOf cellRange Is Range Then
        cellRange = DS_RangeToArray(cellRange)
    End If
    
    If IsMissing(decimals) Then
        decimals = 0
    End If
    
    Dim q1 As Double
    q1 = WorksheetFunction.Quartile(cellRange, 1)
    Dim q3 As Double
    q3 = WorksheetFunction.Quartile(cellRange, 3)
    
    DS_PrintIQR = WorksheetFunction.Round(q1, decimals) & " - " & WorksheetFunction.Round(q3, decimals)
End Function

Public Function DS_MergeRangesToArray(ByVal cellRange1 As Range, Optional cellRange2 As Variant, Optional cellRange3 As Variant, Optional cellRange4 As Variant, Optional cellRange5 As Variant, Optional cellRange6 As Variant, Optional cellRange7 As Variant, Optional cellRange8 As Variant, Optional cellRange9 As Variant, Optional cellRange10 As Variant)
    Dim valueArray() As Variant
    valueArray = DS_RangeToArray(cellRange1)
    
    Dim valueArrayAdd() As Variant
    If Not IsMissing(cellRange2) Then
        valueArrayAdd = DS_RangeToArray(cellRange2)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange3) Then
        valueArrayAdd = DS_RangeToArray(cellRange3)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange4) Then
        valueArrayAdd = DS_RangeToArray(cellRange4)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange5) Then
        valueArrayAdd = DS_RangeToArray(cellRange5)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange6) Then
        valueArrayAdd = DS_RangeToArray(cellRange6)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange7) Then
        valueArrayAdd = DS_RangeToArray(cellRange7)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange8) Then
        valueArrayAdd = DS_RangeToArray(cellRange8)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange9) Then
        valueArrayAdd = DS_RangeToArray(cellRange9)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    If Not IsMissing(cellRange10) Then
        valueArrayAdd = DS_RangeToArray(cellRange10)
        valueArray = DS_AppendToArray(valueArray, valueArrayAdd)
    End If
    
    Dim resultArray As Variant
    ReDim resultArray(0 To 0)
    Dim counter As Integer
    For Each value In valueArray
        If Not IsEmpty(value) Then
            ReDim Preserve resultArray(0 To counter)
            resultArray(counter) = value
            counter = counter + 1
        End If
    Next value
    
    DS_MergeRangesToArray = resultArray
End Function

