Attribute VB_Name = "DSStatTools_ROCCorr_Obuchowski"
Public Function DS_ROCCorr_AUC_Obuchowski(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional isPathologyHigher As Boolean = True) As Double
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim uniqueClusters() As Variant
    Dim presentCases As Variant, absentCases As Variant
    Dim numUniqueClusters As Long
    Dim i As Long, I01 As Long, I10 As Long
    Dim M() As Long, n() As Long
    Dim AUC1 As Double
    Dim Xcomps() As Double, Ycomps() As Double
    Dim sum_X As Double, sum_Y As Double
    Dim totalPositives As Long, totalNegatives As Long
    Dim S10_1 As Double, S01_1 As Double, S11_1 As Double
    Dim var_1 As Double, aucSE As Double, CIlo As Double, CIhi As Double
    
    ' algorithm adapted from R: https://github.com/DIDSR/mitoticFigureCounts/blob/master/R/doAUCcluster.R

    ' Get the arrays from the ranges (for ranges in Excel)
    arrMeasure = DS_RangeToArray(measurementRange)
    arrPathology = DS_RangeToArray(pathologyRange)
    arrCluster = DS_RangeToArray(clusterRange)
    
    ' If isPathologyHigher is False, negate the measurement values
    If Not isPathologyHigher Then
        For i = LBound(arrMeasure) To UBound(arrMeasure)
            arrMeasure(i) = -arrMeasure(i)
        Next i
    End If

    ' Get unique clusters
    uniqueClusters = DS_GetUnique(arrCluster)
    numUniqueClusters = UBound(uniqueClusters) - LBound(uniqueClusters) + 1

    ' Initialize arrays for Xcomps, Ycomps, m, and n
    ReDim Xcomps(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim Ycomps(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim M(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim n(LBound(uniqueClusters) To UBound(uniqueClusters))

    ' Separate the present and absent cases
    presentCases = DS_FilterByPathology(arrMeasure, arrPathology, 1)
    absentCases = DS_FilterByPathology(arrMeasure, arrPathology, 0)

    ' Get the AUC for the current data
    AUC1 = getAUC(presentCases, absentCases)

    ' Return AUC
    DS_ROCCorr_AUC_Obuchowski = AUC1
End Function

Public Function DS_ROCCorr_AUCCI_Obuchowski(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional isPathologyHigher As Boolean = True, Optional significanceLevel As Double = 0.95)
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim uniqueClusters() As Variant
    Dim presentCases As Variant, absentCases As Variant
    Dim numUniqueClusters As Long
    Dim i As Long, I01 As Long, I10 As Long
    Dim M() As Long, n() As Long
    Dim AUC1 As Double
    Dim Xcomps() As Double, Ycomps() As Double
    Dim sum_X As Double, sum_Y As Double
    Dim totalPositives As Long, totalNegatives As Long
    Dim S10_1 As Double, S01_1 As Double, S11_1 As Double
    Dim var_1 As Double, aucSE As Double, CIlo As Double, CIhi As Double

    ' Get the arrays from the ranges (for ranges in Excel)
    arrMeasure = DS_RangeToArray(measurementRange)
    arrPathology = DS_RangeToArray(pathologyRange)
    arrCluster = DS_RangeToArray(clusterRange)
    
    ' If isPathologyHigher is False, negate the measurement values
    If Not isPathologyHigher Then
        For i = LBound(arrMeasure) To UBound(arrMeasure)
            arrMeasure(i) = -arrMeasure(i)
        Next i
    End If

    ' Get unique clusters
    uniqueClusters = DS_GetUnique(arrCluster)
    numUniqueClusters = UBound(uniqueClusters) - LBound(uniqueClusters) + 1

    ' Initialize arrays for Xcomps, Ycomps, m, and n
    ReDim Xcomps(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim Ycomps(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim M(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim n(LBound(uniqueClusters) To UBound(uniqueClusters))

    ' Separate the present and absent cases
    presentCases = DS_FilterByPathology(arrMeasure, arrPathology, 1)
    absentCases = DS_FilterByPathology(arrMeasure, arrPathology, 0)

    ' Get the AUC for the current data
    AUC1 = getAUC(presentCases, absentCases)

    ' Calculate m (number of positives) and n (number of negatives) for each cluster
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        M(i) = DS_CountPosInCluster(uniqueClusters(i), arrCluster, arrPathology) ' Positive cases in cluster
        n(i) = DS_CountNegInCluster(uniqueClusters(i), arrCluster, arrPathology) ' Negative cases in cluster
    Next i

    ' Loop through each cluster to calculate Xcomps and Ycomps
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        ' Calculate Xcomps (V10 sums for positive cases in the cluster)
        If M(i) > 0 Then
            Xcomps(i) = DS_CalculateXcomp(uniqueClusters(i), arrCluster, arrPathology, arrMeasure, 1)
        Else
            Xcomps(i) = 0
        End If

        ' Calculate Ycomps (V01 sums for negative cases in the cluster)
        If n(i) > 0 Then
            Ycomps(i) = DS_CalculateYcomp(uniqueClusters(i), arrCluster, arrPathology, arrMeasure, 0)
        Else
            Ycomps(i) = 0
        End If
    Next i

    ' Calculate total number of positives and negatives across all clusters
    totalPositives = WorksheetFunction.sum(M)
    totalNegatives = WorksheetFunction.sum(n)
    
    I10 = 0
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        If DS_CountPosInCluster(uniqueClusters(i), arrCluster, arrPathology) > 0 Then
            I10 = I10 + 1 ' This cluster has at least one positive case
        End If
    Next i
    
    I01 = 0
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        If DS_CountNegInCluster(uniqueClusters(i), arrCluster, arrPathology) > 0 Then
            I01 = I01 + 1 ' This cluster has at least one negative case
        End If
    Next i

    ' Return CI
    DS_ROCCorr_AUCCI_Obuchowski = DS_ROCCorr_CalculateCI(Xcomps, Ycomps, AUC1, M, n, numUniqueClusters, I10, I01, totalPositives, totalNegatives, significanceLevel)
End Function

Public Function DS_ROCCorr_AUCP_Obuchowski(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional isPathologyHigher As Boolean = True)
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim uniqueClusters() As Variant
    Dim presentCases As Variant, absentCases As Variant
    Dim numUniqueClusters As Long
    Dim i As Long, I01 As Long, I10 As Long
    Dim M() As Long, n() As Long
    Dim AUC1 As Double
    Dim Xcomps() As Double, Ycomps() As Double
    Dim sum_X As Double, sum_Y As Double
    Dim totalPositives As Long, totalNegatives As Long
    Dim S10_1 As Double, S01_1 As Double, S11_1 As Double
    Dim var_1 As Double, aucVar As Double, CIlo As Double, CIhi As Double

    ' Get the arrays from the ranges (for ranges in Excel)
    arrMeasure = DS_RangeToArray(measurementRange)
    arrPathology = DS_RangeToArray(pathologyRange)
    arrCluster = DS_RangeToArray(clusterRange)
    
    ' If isPathologyHigher is False, negate the measurement values
    If Not isPathologyHigher Then
        For i = LBound(arrMeasure) To UBound(arrMeasure)
            arrMeasure(i) = -arrMeasure(i)
        Next i
    End If

    ' Get unique clusters
    uniqueClusters = DS_GetUnique(arrCluster)
    numUniqueClusters = UBound(uniqueClusters) - LBound(uniqueClusters) + 1

    ' Initialize arrays for Xcomps, Ycomps, m, and n
    ReDim Xcomps(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim Ycomps(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim M(LBound(uniqueClusters) To UBound(uniqueClusters))
    ReDim n(LBound(uniqueClusters) To UBound(uniqueClusters))

    ' Separate the present and absent cases
    presentCases = DS_FilterByPathology(arrMeasure, arrPathology, 1)
    absentCases = DS_FilterByPathology(arrMeasure, arrPathology, 0)

    ' Get the AUC for the current data
    AUC1 = getAUC(presentCases, absentCases)

    ' Calculate m (number of positives) and n (number of negatives) for each cluster
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        M(i) = DS_CountPosInCluster(uniqueClusters(i), arrCluster, arrPathology) ' Positive cases in cluster
        n(i) = DS_CountNegInCluster(uniqueClusters(i), arrCluster, arrPathology) ' Negative cases in cluster
    Next i

    ' Loop through each cluster to calculate Xcomps and Ycomps
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        ' Calculate Xcomps (V10 sums for positive cases in the cluster)
        If M(i) > 0 Then
            Xcomps(i) = DS_CalculateXcomp(uniqueClusters(i), arrCluster, arrPathology, arrMeasure, 1)
        Else
            Xcomps(i) = 0
        End If

        ' Calculate Ycomps (V01 sums for negative cases in the cluster)
        If n(i) > 0 Then
            Ycomps(i) = DS_CalculateYcomp(uniqueClusters(i), arrCluster, arrPathology, arrMeasure, 0)
        Else
            Ycomps(i) = 0
        End If
    Next i

    ' Calculate total number of positives and negatives across all clusters
    totalPositives = WorksheetFunction.sum(M)
    totalNegatives = WorksheetFunction.sum(n)
    
    I10 = 0
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        If DS_CountPosInCluster(uniqueClusters(i), arrCluster, arrPathology) > 0 Then
            I10 = I10 + 1 ' This cluster has at least one positive case
        End If
    Next i
    
    I01 = 0
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        If DS_CountNegInCluster(uniqueClusters(i), arrCluster, arrPathology) > 0 Then
            I01 = I01 + 1 ' This cluster has at least one negative case
        End If
    Next i

    
    
    aucVar = DS_ROCCorr_CalculateVar(Xcomps, Ycomps, AUC1, M, n, numUniqueClusters, I10, I01, totalPositives, totalNegatives)
    
    Dim Z As Double
    Dim pValue As Double
    
    If aucVar > 0 Then
        Z = (AUC - 0.5) / Sqr(aucVar)
        ' Step 4: Calculate the p-value from the z-score (two-tailed test)
        pValue = 2 * (1 - Application.WorksheetFunction.NormSDist(Abs(Z)))
    Else
        ' If variance is 0, AUC is perfectly deterministic
        If AUC = 1 Then
            pValue = 0 ' Perfect separation (AUC = 1 or 0)
        Else
            pValue = 1 ' Non-informative test (AUC = 0.5)
        End If
    End If
    
    DS_ROCCorr_AUCP_Obuchowski = pValue
End Function

Private Function DS_ROCCorr_CalculateVar(ByVal Xcomps As Variant, ByVal Ycomps As Variant, ByVal AUC As Double, _
                                       ByVal M As Variant, ByVal n As Variant, _
                                       ByVal numUniqueClusters As Long, ByVal I10 As Long, ByVal I01 As Long, _
                                       ByVal totalPositives As Long, ByVal totalNegatives As Long) As Double
    Dim S10_1 As Double, S01_1 As Double, S11_1 As Double
    Dim var_1 As Double, aucSE As Double, CIlo As Double, CIhi As Double
    Dim i As Long

    ' Initialize sums
    S10_1 = 0
    S01_1 = 0
    S11_1 = 0

    ' Avoid division by zero if there are no positives or negatives
    If totalPositives = 0 Or totalNegatives = 0 Then
        DS_ROCCorr_CalculateSE = Array(0, 1)
        Exit Function
    End If

    ' Calculate sum of squares for Xcomps (S10_1)
    For i = LBound(Xcomps) To UBound(Xcomps)
        S10_1 = S10_1 + (Xcomps(i) - M(i) * AUC) * (Xcomps(i) - M(i) * AUC)
    Next i
    S10_1 = (I10 / ((I10 - 1) * totalPositives)) * S10_1

    ' Calculate sum of squares for Ycomps (S01_1)
    For i = LBound(Ycomps) To UBound(Ycomps)
        S01_1 = S01_1 + (Ycomps(i) - n(i) * AUC) * (Ycomps(i) - n(i) * AUC)
    Next i
    S01_1 = (I01 / ((I01 - 1) * totalNegatives)) * S01_1

    ' Calculate cross-product for Xcomps and Ycomps (S11_1)
    For i = LBound(Xcomps) To UBound(Xcomps)
        S11_1 = S11_1 + (Xcomps(i) - M(i) * AUC) * (Ycomps(i) - n(i) * AUC)
    Next i
    S11_1 = (numUniqueClusters / (numUniqueClusters - 1)) * S11_1

    ' Calculate variance of AUC
    var_1 = S10_1 / totalPositives + S01_1 / totalNegatives + (2 * S11_1) / (totalPositives * totalNegatives)

    ' var of AUC
    DS_ROCCorr_CalculateVar = var_1
End Function


Private Function DS_ROCCorr_CalculateCI(ByVal Xcomps As Variant, ByVal Ycomps As Variant, ByVal AUC As Double, _
                                       ByVal M As Variant, ByVal n As Variant, _
                                       ByVal numUniqueClusters As Long, ByVal I10 As Long, ByVal I01 As Long, _
                                       ByVal totalPositives As Long, ByVal totalNegatives As Long, Optional significanceLevel As Double = 0.95) As Variant
    Dim aucVar As Double, CIlo As Double, CIhi As Double
    Dim Z As Double

    ' Z-score for 95% CI
    Z = WorksheetFunction.Norm_S_Inv(1 - (1 - significanceLevel) / 2)  ' 1.96 for 95% CI
    
    ' Standard error of AUC
    aucVar = DS_ROCCorr_CalculateVar(Xcomps, Ycomps, AUC, M, n, numUniqueClusters, I10, I01, totalPositives, totalNegatives)

    If aucVar <> 0 Then
        ' Confidence Interval (CI)
        CIlo = AUC - Z * Sqr(aucVar)
        CIhi = AUC + Z * Sqr(aucVar)
    
        ' Ensure CI bounds are within [0, 1]
        CIlo = Application.Max(0, CIlo)
        CIhi = Application.Min(1, CIhi)
    
        ' Return CI as an array
        DS_ROCCorr_CalculateCI = Array(CIlo, CIhi)
    End If
End Function

