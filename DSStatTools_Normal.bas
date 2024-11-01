Attribute VB_Name = "DSStatTools_Normal"
Public Function DS_ShapiroWilkP(ByVal cellRange As Range)
    Dim valueArray() As Variant
    valueArray = DS_RangeToArray(cellRange)
    If UBound(valueArray) + 1 < 3 Or UBound(valueArray) + 1 > 2000 Then
        DS_ShapiroWilkP = -1
        Exit Function
    End If
    
    If UBound(valueArray) + 1 < 12 Then
        DS_ShapiroWilkP = DS_ShapiroWilkBasicP(valueArray)
    Else
        DS_ShapiroWilkP = DS_ShapiroWilkExtendedP(valueArray)
    End If
    
End Function

Private Function DS_ShapiroWilkExtendedP(ByVal valueArray As Variant)
    Call DS_QuickSort(valueArray, 0, UBound(valueArray))
    
    Dim avg As Double
    avg = WorksheetFunction.Average(valueArray)
    
    Dim n As Double
    n = UBound(valueArray) + 1
    
    Dim u As Double
    u = 1 / Math.Sqr(n)
    
    Dim M As Double
    
    Dim mArray() As Double
    ReDim mArray(0 To UBound(valueArray))
    
    Dim aArray() As Double
    ReDim aArray(0 To UBound(valueArray))
    
    Dim i As Integer
    
    For i = 0 To n - 1
        mArray(i) = WorksheetFunction.Norm_S_Inv((i + 1 - 0.375) / (n + 0.25))
        M = M + mArray(i) ^ 2
    Next i
    
    aArray(UBound(aArray)) = -2.706056 * u ^ 5 + 4.434685 * u ^ 4 - 2.07119 * u ^ 3 - 0.147981 * u ^ 2 + 0.221157 * u + mArray(UBound(mArray)) / Math.Sqr(M)
    aArray(0) = -aArray(UBound(aArray))
    
    aArray(UBound(aArray) - 1) = -3.582633 * u ^ 5 + 5.682633 * u ^ 4 - 1.752461 * u ^ 3 - 0.293762 * u ^ 2 + 0.042981 * u + mArray(UBound(mArray) - 1) / Math.Sqr(M)
    aArray(1) = -aArray(UBound(aArray) - 1)
    
    Dim epsilon As Double
    epsilon = (M - 2 * mArray(UBound(mArray)) ^ 2 - 2 * mArray(UBound(mArray) - 1) ^ 2) / (1 - 2 * aArray(UBound(aArray)) ^ 2 - 2 * aArray(UBound(aArray) - 1) ^ 2)
    
    For i = 2 To n - 3
        aArray(i) = mArray(i) / Math.Sqr(epsilon)
    Next i
    
    Dim numerator As Double
    numerator = 0 'sum(ai*xi)^2
    Dim denominator As Double
    denominator = 0 ' sum((xi - avg)^2)
    
    For i = 0 To UBound(valueArray)
        numerator = numerator + aArray(i) * valueArray(i)
        denominator = denominator + (valueArray(i) - avg) ^ 2
    Next i
    numerator = numerator * numerator
    
    Dim w As Double
    w = numerator / denominator
    
    Dim my As Double
    my = 0.0038915 * Math.Log(n) ^ 3 - 0.083751 * Math.Log(n) ^ 2 - 0.31082 * Math.Log(n) - 1.5861
    
    Dim sigma As Double
    sd = Math.Exp(0.0030302 * Math.Log(n) ^ 2 - 0.082676 * Math.Log(n) - 0.4803)
    
    Dim Z As Double
    Z = (Math.Log(1 - w) - my) / sd
    
    Dim p As Double
    p = 1 - WorksheetFunction.Norm_S_Dist(Z, True)
    
    DS_ShapiroWilkExtendedP = p
End Function

Private Function DS_ShapiroWilkBasicP(ByVal valueArray As Variant)
    Call DS_QuickSort(valueArray, 0, UBound(valueArray))
    
    Dim avg As Double
    avg = WorksheetFunction.Average(valueArray)
    
    Dim numerator As Double
    numerator = 0 'sum(ai*xi)^2
    Dim denominator As Double
    denominator = 0 ' sum((xi - avg)^2)
    
    Dim weights As Variant
    weights = DS_ShapiroWilkBasicWeights(UBound(valueArray) + 1)
    
    Dim i As Integer
    For i = 0 To UBound(valueArray)
        numerator = numerator + weights(i) * valueArray(i)
        denominator = denominator + (valueArray(i) - avg) ^ 2
    Next i
    numerator = numerator * numerator
    
    Dim w As Double
    w = numerator / denominator
    
    Dim pTable() As Variant
    pTable = DS_ShapiroWilkBasicPTable(UBound(valueArray) + 1)
    
    If w < pTable(0) Then
        DS_ShapiroWilkBasicP = "<0.01"
    ElseIf w > pTable(8) Then
        DS_ShapiroWilkBasicP = ">0.99"
    Else
        For i = 0 To UBound(pTable) - 1
            If w >= pTable(i) And w < pTable(i + 1) Then
                Dim interp As Double
                interp = (w - pTable(i)) / (pTable(i + 1) - pTable(i))
                
                Dim p1 As Double
                Dim p2 As Double
                
                If i = 0 Then
                    p1 = 0.01
                    p2 = 0.02
                ElseIf i = 1 Then
                    p1 = 0.02
                    p2 = 0.05
                ElseIf i = 2 Then
                    p1 = 0.05
                    p2 = 0.1
                ElseIf i = 3 Then
                    p1 = 0.1
                    p2 = 0.5
                ElseIf i = 4 Then
                    p1 = 0.5
                    p2 = 0.9
                ElseIf i = 5 Then
                    p1 = 0.9
                    p2 = 0.95
                ElseIf i = 6 Then
                    p1 = 0.95
                    p2 = 0.98
                ElseIf i = 7 Then
                    p1 = 0.98
                    p2 = 0.99
                End If
                
                DS_ShapiroWilkBasicP = p1 + interp * (p2 - p1)
            End If
        Next i
    End If
End Function

Private Function DS_ShapiroWilkBasicWeights(ByVal num As Integer)
    Dim weights() As Variant

    If num = 2 Then
        weights = Array(0.7071)
    ElseIf num = 3 Then
        weights = Array(0.7071)
    ElseIf num = 4 Then
        weights = Array(0.6872, 0.1677)
    ElseIf num = 5 Then
        weights = Array(0.6646, 0.2413)
    ElseIf num = 6 Then
        weights = Array(0.6431, 0.2806, 0.0875)
    ElseIf num = 7 Then
        weights = Array(0.6233, 0.3031, 0.1401)
    ElseIf num = 8 Then
        weights = Array(0.6052, 0.3164, 0.1743, 0.0561)
    ElseIf num = 9 Then
        weights = Array(0.5888, 0.3244, 0.1976, 0.0947)
    ElseIf num = 10 Then
        weights = Array(0.5739, 0.3291, 0.214, 0.1224, 0.0399)
    ElseIf num = 11 Then
        weights = Array(0.5601, 0.3315, 0.226, 0.1429, 0.0695)
    End If
    
    ReDim Preserve weights(0 To num - 1)
    
    If num Mod 2 = 1 Then
        weights(WorksheetFunction.RoundDown(num / 2, 0)) = 0
    End If
    
    Dim i As Integer
    For i = 0 To WorksheetFunction.RoundDown(num / 2, 0) - 1
        weights(num - i - 1) = -weights(i)
    Next i
    
    DS_ShapiroWilkBasicWeights = weights
End Function

Private Function DS_ShapiroWilkBasicPTable(ByVal num As Integer)
    If num = 3 Then
        DS_ShapiroWilkBasicPTable = Array(0.753, 0.756, 0.767, 0.789, 0.959, 0.998, 0.999, 1, 1)
    ElseIf num = 4 Then
        DS_ShapiroWilkBasicPTable = Array(0.687, 0.707, 0.748, 0.792, 0.935, 0.987, 0.992, 0.996, 0.997)
    ElseIf num = 5 Then
        DS_ShapiroWilkBasicPTable = Array(0.686, 0.715, 0.762, 0.806, 0.927, 0.979, 0.986, 0.991, 0.993)
    ElseIf num = 6 Then
        DS_ShapiroWilkBasicPTable = Array(0.713, 0.743, 0.788, 0.826, 0.927, 0.974, 0.981, 0.986, 0.989)
    ElseIf num = 7 Then
        DS_ShapiroWilkBasicPTable = Array(0.73, 0.76, 0.803, 0.838, 0.928, 0.972, 0.979, 0.985, 0.988)
    ElseIf num = 8 Then
        DS_ShapiroWilkBasicPTable = Array(0.749, 0.778, 0.818, 0.851, 0.932, 0.972, 0.978, 0.984, 0.987)
    ElseIf num = 9 Then
        DS_ShapiroWilkBasicPTable = Array(0.764, 0.791, 0.829, 0.859, 0.935, 0.972, 0.978, 0.984, 0.986)
    ElseIf num = 10 Then
        DS_ShapiroWilkBasicPTable = Array(0.781, 0.806, 0.842, 0.869, 0.938, 0.972, 0.978, 0.983, 0.986)
    ElseIf num = 11 Then
        DS_ShapiroWilkBasicPTable = Array(0.792, 0.817, 0.85, 0.876, 0.94, 0.973, 0.979, 0.984, 0.986)
    End If
End Function
