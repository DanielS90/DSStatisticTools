Public Function DS_UTestP(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional sided As Variant)
    If TypeOf cellRange1 Is Range Then
        cellRange1 = DS_RangeToArray(cellRange1)
    End If
    
    If TypeOf cellRange2 Is Range Then
        cellRange2 = DS_RangeToArray(cellRange2)
    End If
    
    Dim allValues() As Variant
    allValues = DS_JoinArrays(cellRange1, cellRange2)
    
    If IsMissing(sided) Then
        sided = 2
    End If
    
    Dim n1 As Double
    n1 = UBound(cellRange1) - LBound(cellRange1) + 1
    
    Dim n2 As Double
    n2 = UBound(cellRange2) - LBound(cellRange2) + 1
    
    Dim ranks1() As Double
    ReDim ranks1(LBound(cellRange1) To UBound(cellRange1))
    
    Dim rankSum1 As Double
    
    Dim ranks2() As Double
    ReDim ranks2(LBound(cellRange2) To UBound(cellRange2))
    
    Dim rankSum2 As Double
    
    Dim i As Integer
    For i = LBound(cellRange1) To UBound(cellRange1)
        ranks1(i) = DS_Rank(cellRange1(i), allValues)
        rankSum1 = rankSum1 + ranks1(i)
    Next i
    
    For i = LBound(cellRange2) To UBound(cellRange2)
        ranks2(i) = DS_Rank(cellRange2(i), allValues)
        rankSum2 = rankSum2 + ranks2(i)
    Next i
    
    Dim t As Double
    Dim ns As Double
    Dim nl As Double
    
    If rankSum1 < rankSum2 Then
        t = rankSum1
        ns = n1
        nl = n2
    Else
        t = rankSum2
        ns = n2
        nl = n1
    End If
    
    Dim my As Double
    my = ns * (ns + nl + 1) / 2
    
    Dim sd As Double
    sd = Math.Sqr(nl * my / 6)
    
    Dim z As Double
    z = (t - my) / sd
    
    DS_UTestP = WorksheetFunction.Norm_S_Dist(z, True)
    If z > 0 Then
        DS_UTestP = 1 - DS_UTestP
    End If
    
    If sided = 2 Then
        DS_UTestP = 2 * DS_UTestP
    End If
End Function

Public Function DS_Rank(ByVal value As Variant, ByVal allValues As Variant)
    Call DS_QuickSort(allValues, LBound(allValues), UBound(allValues))
    
    Dim rankSum As Double
    Dim occurrences As Integer
    
    Dim i As Integer
    For i = LBound(allValues) To UBound(allValues)
        If allValues(i) = value Then
            rankSum = rankSum + i - LBound(allValues) + 1
            occurrences = occurrences + 1
        ElseIf occurrences > 0 Then
            Exit For
        End If
    Next i
    
    If occurrences > 0 Then
        DS_Rank = rankSum / occurrences
    End If
End Function
