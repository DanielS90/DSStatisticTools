Public Function DS_Correlation_PointBiserialR(ByVal metricRange As Variant, ByVal binaryRange As Variant)
    If TypeOf metricRange Is Range Then
        metricRange = DS_RangeToArray(metricRange)
    End If
    If TypeOf binaryRange Is Range Then
        binaryRange = DS_RangeToArray(binaryRange)
    End If
    
    If Not UBound(metricRange) = UBound(binaryRange) Then
        Exit Function
    End If
    
    Dim count As Integer
    count = UBound(metricRange) - LBound(metricRange) + 1
    
    Dim binaryValues() As Variant
    binaryValues = DS_UniqueValues(binaryRange)
    
    If IsEmpty(binaryValues) Or Not UBound(binaryValues) = 1 Then
        Exit Function
    End If
    
    Dim n1 As Integer
    Dim n2 As Integer
    Dim avgY1 As Double
    Dim avgY2 As Double
    
    Dim currentIndex As Integer
    For currentIndex = LBound(metricRange) To UBound(metricRange)
        If binaryRange(currentIndex) = binaryValues(LBound(binaryValues)) Then
            n1 = n1 + 1
            avgY1 = avgY1 + metricRange(currentIndex)
        Else
            n2 = n2 + 1
            avgY2 = avgY2 + metricRange(currentIndex)
        End If
    Next currentIndex
    avgY1 = avgY1 / n1
    avgY2 = avgY2 / n2
    
    Dim p As Double
    p = n1 / count
    
    Dim q As Double
    q = n2 / count
    
    Dim sd As Double
    sd = WorksheetFunction.StDev_P(metricRange)
    
    DS_Correlation_PointBiserialR = (avgY1 - avgY2) * Math.Sqr(p * q) / sd
    
End Function

Public Function DS_Correlation_PointBiserialP(ByVal metricRange As Variant, ByVal binaryRange As Variant)
    If TypeOf metricRange Is Range Then
        metricRange = DS_RangeToArray(metricRange)
    End If
    If TypeOf binaryRange Is Range Then
        binaryRange = DS_RangeToArray(binaryRange)
    End If
    
    Dim r As Double
    r = DS_Correlation_PointBiserialR(metricRange, binaryRange)
    
    Dim count As Integer
    count = UBound(metricRange) - LBound(metricRange) + 1
    
    Dim t As Double
    t = r * Math.Sqr(count - 2) / Math.Sqr(1 - r ^ 2)
    
    If t < 0 Then
        t = t * -1
    End If
    
    DS_Correlation_PointBiserialP = WorksheetFunction.T_Dist_2T(t, count - 2)
End Function

Public Function DS_Correlation_SpearmanR(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant)
    If TypeOf cellRange1 Is Range Then
        cellRange1 = DS_RangeToArray(cellRange1)
    End If
    If TypeOf cellRange2 Is Range Then
        cellRange2 = DS_RangeToArray(cellRange2)
    End If
    
    If Not UBound(cellRange1) = UBound(cellRange2) Then
        Exit Function
    End If
    
    Dim ranks1() As Variant
    ReDim ranks1(LBound(cellRange1) To UBound(cellRange1))
    
    Dim ranks2() As Variant
    ReDim ranks2(LBound(cellRange2) To UBound(cellRange2))
    
    Dim currentIndex As Integer
    For currentIndex = LBound(cellRange1) To UBound(cellRange1)
        ranks1(currentIndex) = DS_Rank(cellRange1(currentIndex), cellRange1)
        ranks2(currentIndex) = DS_Rank(cellRange2(currentIndex), cellRange2)
    Next currentIndex
    
    Dim r As Double
    r = WorksheetFunction.Correl(ranks1, ranks2)
    
    DS_Correlation_SpearmanR = r
End Function

Public Function DS_Correlation_SpearmanP(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant)
    If TypeOf cellRange1 Is Range Then
        cellRange1 = DS_RangeToArray(cellRange1)
    End If
    If TypeOf cellRange2 Is Range Then
        cellRange2 = DS_RangeToArray(cellRange2)
    End If
    
    If Not UBound(cellRange1) = UBound(cellRange2) Then
        Exit Function
    End If
    
    Dim r As Double
    r = DS_Correlation_SpearmanR(cellRange1, cellRange2)
    
    Dim n As Integer
    n = UBound(cellRange1) - LBound(cellRange1) + 1
    
    Dim t As Double
    t = Abs(r) * Math.Sqr(n - 2) / Math.Sqr(1 - r * r)
    
    Dim df As Integer
    df = n - 2
    
    Dim p As Double
    p = WorksheetFunction.T_Dist_2T(t, df)
    
    DS_Correlation_SpearmanP = p
End Function

Public Function DS_Correlation_Spearman95CI(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional decimals As Variant)
    If TypeOf cellRange1 Is Range Then
        cellRange1 = DS_RangeToArray(cellRange1)
    End If
    If TypeOf cellRange2 Is Range Then
        cellRange2 = DS_RangeToArray(cellRange2)
    End If
    
    If Not UBound(cellRange1) = UBound(cellRange2) Then
        Exit Function
    End If
    
    If IsMissing(decimals) Then
        decimals = 2
    End If
    
    Dim r As Double
    r = DS_Correlation_SpearmanR(cellRange1, cellRange2)
    
    Dim n As Integer
    n = UBound(cellRange1) - LBound(cellRange1) + 1
    
    Dim sd As Double
    sd = 1 / Math.Sqr(n - 3)
    
    Dim lower As Double
    lower = WorksheetFunction.Tanh(WorksheetFunction.Atanh(r) - 1.96 * sd)
    
    Dim upper As Double
    upper = WorksheetFunction.Tanh(WorksheetFunction.Atanh(r) + 1.96 * sd)
    
    DS_Correlation_Spearman95CI = Math.Round(lower, decimals) & " - " & Math.Round(upper, decimals)
End Function

Public Function DS_Correlation_PearsonR(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant)
    If TypeOf cellRange1 Is Range Then
        cellRange1 = DS_RangeToArray(cellRange1)
    End If
    If TypeOf cellRange2 Is Range Then
        cellRange2 = DS_RangeToArray(cellRange2)
    End If
    
    If Not UBound(cellRange1) = UBound(cellRange2) Then
        Exit Function
    End If
    
    Dim r As Double
    r = WorksheetFunction.Pearson(cellRange1, cellRange2)
    
    DS_Correlation_PearsonR = r
End Function

Public Function DS_Correlation_PearsonP(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant)
    If TypeOf cellRange1 Is Range Then
        cellRange1 = DS_RangeToArray(cellRange1)
    End If
    If TypeOf cellRange2 Is Range Then
        cellRange2 = DS_RangeToArray(cellRange2)
    End If
    
    If Not UBound(cellRange1) = UBound(cellRange2) Then
        Exit Function
    End If
    
    Dim r As Double
    r = DS_Correlation_PearsonR(cellRange1, cellRange2)
    
    Dim n As Integer
    n = UBound(cellRange1) - LBound(cellRange1) + 1
    
    Dim t As Double
    t = Abs(r) * Math.Sqr(n - 2) / Math.Sqr(1 - r * r)
    
    Dim df As Integer
    df = n - 2
    
    Dim p As Double
    p = WorksheetFunction.T_Dist_2T(t, df)
    
    DS_Correlation_PearsonP = p
End Function

Public Function DS_Correlation_Pearson95CI(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional decimals As Variant)
    If TypeOf cellRange1 Is Range Then
        cellRange1 = DS_RangeToArray(cellRange1)
    End If
    If TypeOf cellRange2 Is Range Then
        cellRange2 = DS_RangeToArray(cellRange2)
    End If
    
    If Not UBound(cellRange1) = UBound(cellRange2) Then
        Exit Function
    End If
    
    If IsMissing(decimals) Then
        decimals = 2
    End If
    
    Dim r As Double
    r = DS_Correlation_PearsonR(cellRange1, cellRange2)
    
    Dim n As Integer
    n = UBound(cellRange1) - LBound(cellRange1) + 1
    
    Dim sd As Double
    sd = 1 / Math.Sqr(n - 3)
    
    Dim lower As Double
    lower = WorksheetFunction.Tanh(WorksheetFunction.Atanh(r) - 1.96 * sd)
    
    Dim upper As Double
    upper = WorksheetFunction.Tanh(WorksheetFunction.Atanh(r) + 1.96 * sd)
    
    DS_Correlation_Pearson95CI = Math.Round(lower, decimals) & " - " & Math.Round(upper, decimals)
End Function
