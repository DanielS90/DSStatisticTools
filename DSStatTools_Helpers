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
