Attribute VB_Name = "DSStatTools_Helpers"
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

Public Function DS_ArrayInitialized(ByVal myArray As Variant) As Boolean
    On Error Resume Next
    If Not IsArray(myArray) Then Exit Function
    Dim i As Integer
    i = UBound(myArray)
    DS_ArrayInitialized = Err.Number = 0
End Function

Public Function DS_JoinArrays(ByVal array1 As Variant, ByVal array2 As Variant)
    Dim n1 As Double
    If DS_ArrayInitialized(array1) Then
        n1 = UBound(array1) - LBound(array1) + 1
    Else
        n1 = 0
    End If
    
    Dim n2 As Double
    If DS_ArrayInitialized(array2) Then
        n2 = UBound(array2) - LBound(array2) + 1
    Else
        n2 = 0
    End If
    
    Dim result() As Variant
    If n1 = 0 And n2 = 0 Then
        result = array1
        DS_JoinArrays = result
        Exit Function
    End If
    
    ReDim result(0 To n1 + n2 - 1)
    
    Dim counter As Integer
    
    Dim val As Variant
    If n1 > 0 Then
        For Each val In array1
            result(counter) = val
            counter = counter + 1
        Next val
    End If
        
    If n2 > 0 Then
        For Each val In array2
            result(counter) = val
            counter = counter + 1
        Next val
    End If
    
    DS_JoinArrays = result
End Function

Public Function DS_AppendToArray(ByVal myArray As Variant, ByVal value As Variant)
    Dim result() As Variant
    Dim counter As Integer
    Dim val As Variant
    
    Dim n1 As Integer
    If DS_ArrayInitialized(myArray) Then
        n1 = UBound(myArray) - LBound(myArray) + 1
    Else
        n1 = 0
    End If
    
    Dim n2 As Integer
    
    If IsArray(value) Then
        n2 = UBound(value) - LBound(value) + 1
        ReDim result(0 To n1 + n2 - 1)
        
        If DS_ArrayInitialized(myArray) Then
            For Each val In myArray
                result(counter) = val
                counter = counter + 1
            Next val
        End If
        
        For Each val In value
            result(counter) = val
            counter = counter + 1
        Next val
    Else
        n2 = 1
        ReDim result(0 To n1 + n2 - 1)
        
        If DS_ArrayInitialized(myArray) Then
            For Each val In myArray
                result(counter) = val
                counter = counter + 1
            Next val
        End If
        
        result(counter) = value
        counter = counter + 1
    End If
    
    DS_AppendToArray = result
End Function

Public Function DS_InArray(ByVal myArray As Variant, ByVal myValue As Variant)
    Dim current As Variant
    For Each current In myArray
        If current Like myValue Then
            DS_InArray = True
            Exit Function
        End If
    Next current
    DS_InArray = False
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

Function DS_ValueInArray(needle As Variant, haystack As Variant) As Boolean
  DS_ValueInArray = (UBound(Filter(haystack, needle)) > -1)
End Function

Function DS_OffsetValues(ByVal indexes As Variant, ByVal reach As Integer, ByVal value1 As Variant, ByVal value2 As Variant)
    If TypeOf indexes Is Range Then
        indexes = DS_RangeToArray(indexes)
    End If
    
    Dim allIndexes() As Variant
    
    Dim currentVal As Variant
    For Each currentVal In indexes
        If InStr(currentVal, ",") > 0 Then
            allIndexes = DS_AppendToArray(allIndexes, DS_CSVStringAsArray(currentVal))
        Else
            allIndexes = DS_AppendToArray(allIndexes, currentVal)
        End If
    Next currentVal
    
    Dim result() As Variant
    ReDim result(0 To reach - 1)
    
    Dim counter As Integer
    For counter = 0 To reach - 1
        If DS_InArray(allIndexes, counter) Then
            result(counter) = value1
        Else
            result(counter) = value2
        End If
    Next counter
    
    DS_OffsetValues = result
End Function

Function DS_FirstOrDefault(ByVal val1 As Variant, Optional val2 As Variant, Optional val3 As Variant, Optional val4 As Variant, Optional val5 As Variant, Optional val6 As Variant, Optional val7 As Variant, Optional val8 As Variant, Optional val9 As Variant, Optional val10 As Variant)
    If IsArray(val1) Then
        If UBound(val1) - LBound(val1) + 1 > 1 Or Not IsEmpty(val1(LBound(val1))) Then
            DS_FirstOrDefault = val1
            Exit Function
        End If
    Else
        If Not IsEmpty(val1) And Not val1 = 0 Then
            DS_FirstOrDefault = val1
            Exit Function
        End If
    End If
    
    If Not IsMissing(val2) Then
        If IsArray(val2) Then
            If UBound(val2) - LBound(val2) + 1 > 1 Or Not IsEmpty(val2(LBound(val2))) Then
                DS_FirstOrDefault = val2
                Exit Function
            End If
        Else
            If Not IsEmpty(val2) And Not val2 = 0 Then
                DS_FirstOrDefault = val2
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val3) Then
        If IsArray(val3) Then
            If UBound(val3) - LBound(val3) + 1 > 1 Or Not IsEmpty(val3(LBound(val3))) Then
                DS_FirstOrDefault = val3
                Exit Function
            End If
        Else
            If Not IsEmpty(val3) And Not val3 = 0 Then
                DS_FirstOrDefault = val3
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val4) Then
        If IsArray(val4) Then
            If UBound(val4) - LBound(val4) + 1 > 1 Or Not IsEmpty(val4(LBound(val4))) Then
                DS_FirstOrDefault = val4
                Exit Function
            End If
        Else
            If Not IsEmpty(val4) And Not val4 = 0 Then
                DS_FirstOrDefault = val4
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val5) Then
        If IsArray(val5) Then
            If UBound(val5) - LBound(val5) + 1 > 1 Or Not IsEmpty(val5(LBound(val5))) Then
                DS_FirstOrDefault = val5
                Exit Function
            End If
        Else
            If Not IsEmpty(val5) And Not val5 = 0 Then
                DS_FirstOrDefault = val5
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val6) Then
        If IsArray(val6) Then
            If UBound(val6) - LBound(val6) + 1 > 1 Or Not IsEmpty(val6(LBound(val6))) Then
                DS_FirstOrDefault = val6
                Exit Function
            End If
        Else
            If Not IsEmpty(val6) And Not val6 = 0 Then
                DS_FirstOrDefault = val6
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val7) Then
        If IsArray(val7) Then
            If UBound(val7) - LBound(val7) + 1 > 1 Or Not IsEmpty(val7(LBound(val7))) Then
                DS_FirstOrDefault = val7
                Exit Function
            End If
        Else
            If Not IsEmpty(val7) And Not val7 = 0 Then
                DS_FirstOrDefault = val7
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val8) Then
        If IsArray(val8) Then
            If UBound(val8) - LBound(val8) + 1 > 1 Or Not IsEmpty(val8(LBound(val8))) Then
                DS_FirstOrDefault = val8
                Exit Function
            End If
        Else
            If Not IsEmpty(val8) And Not val8 = 0 Then
                DS_FirstOrDefault = val8
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val9) Then
        If IsArray(val9) Then
            If UBound(val9) - LBound(val9) + 1 > 1 Or Not IsEmpty(val9(LBound(val9))) Then
                DS_FirstOrDefault = val9
                Exit Function
            End If
        Else
            If Not IsEmpty(val9) And Not val9 = 0 Then
                DS_FirstOrDefault = val9
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(val10) Then
        If IsArray(val10) Then
            If UBound(val10) - LBound(val10) + 1 > 1 Or Not IsEmpty(val10(LBound(val10))) Then
                DS_FirstOrDefault = val10
                Exit Function
            End If
        Else
            If Not IsEmpty(val10) And Not val10 = 0 Then
                DS_FirstOrDefault = val10
                Exit Function
            End If
        End If
    End If
    
End Function

