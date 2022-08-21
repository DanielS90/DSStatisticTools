Option Compare Text

Public Function DS_StringToNumber(ByVal value As Variant, Optional decimalSeparator As Variant, Optional thousandsSeparator As Variant)
    If Not IsMissing(thousandsSeparator) Then
        value = Replace(value, thousandsSeparator, "")
    End If

    If Not IsMissing(decimalSeparator) Then
        value = Replace(value, decimalSeparator, Application.decimalSeparator)
    End If

    DS_StringToNumber = CDbl(value)
End Function

Public Function DS_StringCountOccurrences(ByVal strText As String, ByVal strFind As String) As Long
    Dim lngPos As Long
    Dim lngTemp As Long
    Dim lngCount As Long
    If Len(strText) = 0 Then Exit Function
    If Len(strFind) = 0 Then Exit Function
    lngPos = 1
    Do
        lngPos = InStr(lngPos, strText, strFind)
        lngTemp = lngPos
        If lngPos > 0 Then
            lngCount = lngCount + 1
            lngPos = lngPos + Len(strFind)
        End If
    Loop Until lngPos = 0
    DS_StringCountOccurrences = lngCount
End Function

Public Function DS_IsNumeric(Optional value As Variant)
    If IsMissing(value) Then
        DS_IsNumeric = False
        Exit Function
    End If

    If IsNumeric(value) Then
        Dim numThousandSeparators As Integer
        numThousandSeparators = DS_StringCountOccurrences(value, Application.thousandsSeparator)
        
        If numThousandSeparators > 0 Then
            Dim values() As String
            Dim counter As Integer
            counter = 0
            values = split(value, Application.thousandsSeparator)
            Dim val As Variant
            For Each val In values
                If counter = UBound(values) Then
                    val = split(val, Application.decimalSeparator)(0)
                End If
                
                If counter > 0 And Not Len(val) = 3 Then
                    DS_IsNumeric = False
                    Exit Function
                End If
                counter = counter + 1
            Next val
        End If
        
        DS_IsNumeric = True
    Else
        DS_IsNumeric = False
    End If
End Function

Public Sub DS_QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then DS_QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then DS_QuickSort vArray, tmpLow, inHi
End Sub

Public Function DS_JoinArrays(ByVal array1 As Variant, ByVal array2 As Variant)
    Dim n1 As Double
    n1 = UBound(array1) - LBound(array1) + 1
    
    Dim n2 As Double
    n2 = UBound(array2) - LBound(array2) + 1
    
    Dim result() As Variant
    ReDim result(0 To n1 + n2 - 1)
    
    Dim counter As Integer
    
    Dim val As Variant
    For Each val In array1
        result(counter) = val
        counter = counter + 1
    Next val
    
    For Each val In array2
        result(counter) = val
        counter = counter + 1
    Next val
    
    DS_JoinArrays = result
End Function

Public Function DS_Occurrences(ByVal cellRange As Variant, ByVal comp As Variant, Optional comp2 As Variant, Optional comp3 As Variant, Optional comp4 As Variant)
    If TypeOf cellRange Is Range Then
        cellRange = DS_RangeToArray(cellRange)
    End If
    
    Dim result As Integer
    
    Dim val As Variant
    For Each val In cellRange
        If DS_PatternMatch(val, comp) Then
            result = result + 1
        ElseIf DS_PatternMatch(val, comp2) Then
            result = result + 1
        ElseIf DS_PatternMatch(val, comp3) Then
            result = result + 1
        ElseIf DS_PatternMatch(val, comp4) Then
            result = result + 1
        End If
    Next val
    
    DS_Occurrences = result
End Function

Public Function DS_OccurrencesNot(ByVal cellRange As Variant, ByVal comp As Variant, Optional andNot2 As Variant, Optional andNot3 As Variant, Optional andNot4 As Variant)
    If TypeOf cellRange Is Range Then
        cellRange = DS_RangeToArray(cellRange)
    End If
    
    Dim result As Integer
    
    Dim val As Variant
    For Each val In cellRange
        Dim match As Boolean
        match = DS_PatternMatch(val, comp)
        
        If Not IsMissing(andNot2) Then
            If DS_PatternMatch(val, andNot2) Then
                match = True
            End If
        End If
        
        If Not IsMissing(andNot3) Then
            If DS_PatternMatch(val, andNot3) Then
                match = True
            End If
        End If
        
        If Not IsMissing(andNot4) Then
            If DS_PatternMatch(val, andNot4) Then
                match = True
            End If
        End If
    
        If Not match Then
            result = result + 1
        End If
    Next val
    
    DS_OccurrencesNot = result
End Function

Public Function DS_PatternMatch(ByVal val As Variant, ByVal comp As Variant)
    DS_PatternMatch = False
    If IsMissing(val) Or IsMissing(comp) Then
        Exit Function
    End If
    
    If comp Like "<=*" Then
        If val <= DS_StringToNumber(Right(comp, Len(comp) - 2), ".") Then
            DS_PatternMatch = True
        End If
    ElseIf comp Like "<*" Then
        If val < DS_StringToNumber(Right(comp, Len(comp) - 1), ".") Then
            DS_PatternMatch = True
        End If
    ElseIf comp Like ">=*" Then
        If val >= DS_StringToNumber(Right(comp, Len(comp) - 2), ".") Then
            DS_PatternMatch = True
        End If
    ElseIf comp Like ">*" Then
        If val > DS_StringToNumber(Right(comp, Len(comp) - 1), ".") Then
            DS_PatternMatch = True
        End If
    Else
        If val Like comp Then
            DS_PatternMatch = True
        End If
    End If
End Function

Public Function DS_Max(ByVal cellRange As Variant)
    If TypeOf cellRange Is Range Then
        cellRange = DS_RangeToArray(cellRange)
    End If
    
    Dim result As Variant
    result = cellRange(LBound(cellRange))
    
    Dim val As Variant
    For Each val In cellRange
        If val > result Then
            result = val
        End If
    Next val
    
    DS_Max = result
End Function

Public Function DS_Min(ByVal cellRange As Variant)
    If TypeOf cellRange Is Range Then
        cellRange = DS_RangeToArray(cellRange)
    End If
    
    Dim result As Variant
    result = cellRange(LBound(cellRange))
    
    Dim val As Variant
    For Each val In cellRange
        If val < result Then
            result = val
        End If
    Next val
    
    DS_Min = result
End Function
