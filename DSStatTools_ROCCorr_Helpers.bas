Attribute VB_Name = "DSStatTools_ROCCorr_Helpers"
Function kern(ByVal x1 As Double, ByVal x0 As Double) As Double
    If x1 > x0 Then
        kern = 1
    ElseIf x1 = x0 Then
        kern = 0.5
    Else
        kern = 0
    End If
End Function

Function V10(ByVal Xi As Double, ByVal Ys As Variant) As Double
    Dim sum As Double
    Dim i As Long
    Dim lenYs As Long

    sum = 0
    lenYs = UBound(Ys) - LBound(Ys) + 1

    For i = LBound(Ys) To UBound(Ys)
        sum = sum + kern(Xi, Ys(i))
    Next i

    V10 = sum / lenYs
End Function

Function V01(ByVal Xs As Variant, ByVal Yi As Double) As Double
    Dim sum As Double
    Dim i As Long
    Dim lenXs As Long

    sum = 0
    lenXs = UBound(Xs) - LBound(Xs) + 1

    For i = LBound(Xs) To UBound(Xs)
        sum = sum + kern(Xs(i), Yi)
    Next i

    V01 = sum / lenXs
End Function

Function getAUC(ByVal Xs As Variant, ByVal Ys As Variant) As Double
    Dim sum As Double
    Dim i As Long, j As Long
    Dim lenXs As Long, lenYs As Long

    sum = 0
    lenXs = UBound(Xs) - LBound(Xs) + 1
    lenYs = UBound(Ys) - LBound(Ys) + 1

    For i = LBound(Xs) To UBound(Xs)
        For j = LBound(Ys) To UBound(Ys)
            sum = sum + kern(Xs(i), Ys(j))
        Next j
    Next i

    getAUC = sum / (lenXs * lenYs)
End Function

Function DS_FilterByPathology(measurements As Variant, pathology As Variant, pathologyValue As Integer)
    Dim filtered() As Double
    Dim count As Long
    Dim i As Long

    count = 0
    For i = LBound(pathology) To UBound(pathology)
        If pathology(i) = pathologyValue Then
            count = count + 1
            ReDim Preserve filtered(1 To count)
            filtered(count) = measurements(i)
        End If
    Next i

    DS_FilterByPathology = filtered
End Function

Function DS_CountPosInCluster(clusterID As Variant, clusterRange As Variant, pathologyRange As Variant) As Long
    Dim count As Long
    Dim i As Long

    count = 0
    For i = LBound(clusterRange) To UBound(clusterRange)
        If clusterRange(i) = clusterID And pathologyRange(i) = 1 Then
            count = count + 1
        End If
    Next i

    DS_CountPosInCluster = count
End Function

Function DS_CalculateXcomp(clusterID As Variant, clusterRange As Variant, pathologyRange As Variant, measurements As Variant, pathologyValue As Integer) As Double
    Dim sumXcomp As Double
    Dim i As Long
    Dim presentCases() As Double
    Dim absentCases() As Double

    sumXcomp = 0
    presentCases = DS_FilterByPathology(measurements, pathologyRange, 1)
    absentCases = DS_FilterByPathology(measurements, pathologyRange, 0)

    For i = LBound(clusterRange) To UBound(clusterRange)
        If clusterRange(i) = clusterID And pathologyRange(i) = pathologyValue Then
            sumXcomp = sumXcomp + V10(measurements(i), absentCases)
        End If
    Next i

    DS_CalculateXcomp = sumXcomp
End Function

Function DS_CountNegInCluster(clusterID As Variant, clusterRange As Variant, pathologyRange As Variant) As Long
    Dim count As Long
    Dim i As Long

    count = 0
    For i = LBound(clusterRange) To UBound(clusterRange)
        If clusterRange(i) = clusterID And pathologyRange(i) = 0 Then
            count = count + 1
        End If
    Next i

    DS_CountNegInCluster = count
End Function

Function DS_CalculateYcomp(clusterID As Variant, clusterRange As Variant, pathologyRange As Variant, measurements As Variant, pathologyValue As Integer) As Double
    Dim sumYcomp As Double
    Dim i As Long
    Dim presentCases() As Double
    Dim absentCases() As Double

    sumYcomp = 0
    presentCases = DS_FilterByPathology(measurements, pathologyRange, 1)
    absentCases = DS_FilterByPathology(measurements, pathologyRange, 0)

    For i = LBound(clusterRange) To UBound(clusterRange)
        If clusterRange(i) = clusterID And pathologyRange(i) = pathologyValue Then
            sumYcomp = sumYcomp + V01(presentCases, measurements(i))
        End If
    Next i

    DS_CalculateYcomp = sumYcomp
End Function

Function DS_CalculateS10(Xcomps As Variant, M As Variant, AUC As Double) As Double
    Dim S10 As Double
    Dim i As Long

    S10 = 0
    For i = LBound(Xcomps) To UBound(Xcomps)
        S10 = S10 + (Xcomps(i) - M(i) * AUC) * (Xcomps(i) - M(i) * AUC)
    Next i

    DS_CalculateS10 = S10
End Function

Function DS_CalculateS01(Ycomps As Variant, n As Variant, AUC As Double) As Double
    Dim S01 As Double
    Dim i As Long

    S01 = 0
    For i = LBound(Ycomps) To UBound(Ycomps)
        S01 = S01 + (Ycomps(i) - n(i) * AUC) * (Ycomps(i) - n(i) * AUC)
    Next i

    DS_CalculateS01 = S01
End Function

Function DS_CalculateS11(Xcomps As Variant, Ycomps As Variant, M As Variant, n As Variant, AUC As Double) As Double
    Dim S11 As Double
    Dim i As Long

    S11 = 0
    For i = LBound(Xcomps) To UBound(Xcomps)
        S11 = S11 + (Xcomps(i) - M(i) * AUC) * (Ycomps(i) - n(i) * AUC)
    Next i

    DS_CalculateS11 = S11
End Function
