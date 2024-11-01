Attribute VB_Name = "DSStatTools_ROCCorr_Bootstrap"
Public Function InternalDS_CalculateOptimalCutoff(values As Variant, pathologies As Variant, clusters As Variant, isPathologyHigher As Boolean, ParamArray extraParams() As Variant) As Double
    InternalDS_CalculateOptimalCutoff = DS_ROC_OptimalCutoffYouden(DS_Filter(values, pathologies, 1), DS_Filter(values, pathologies, 0), isPathologyHigher, 100)
End Function

Public Function DS_ROCCorr_OptimalCutoff_Bootstrap(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional isPathologyHigher As Boolean = True, Optional useBCACorrection As Boolean = True, Optional numBootstrap As Integer = 500) As Double()
    DS_ROCCorr_OptimalCutoff_Bootstrap = DS_ROCCorr_Bootstrap_Generic(measurementRange, pathologyRange, clusterRange, "InternalDS_CalculateOptimalCutoff", isPathologyHigher, useBCACorrection, numBootstrap)
End Function

Public Function InternalDS_CalculateSensitivity(values As Variant, pathologies As Variant, clusters As Variant, isPathologyHigher As Boolean, ParamArray extraParams() As Variant) As Double
    InternalDS_CalculateSensitivity = DS_ROC_Sensitivity(DS_Filter(values, pathologies, 1), DS_Filter(values, pathologies, 0), extraParams(0)(0), isPathologyHigher)
End Function

Public Function DS_ROCCorr_Sensitivity_Bootstrap(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional useBCACorrection As Boolean = True, Optional numBootstrap As Integer = 500) As Double()
    DS_ROCCorr_Sensitivity_Bootstrap = DS_ROCCorr_Bootstrap_Generic(measurementRange, pathologyRange, clusterRange, "InternalDS_CalculateSensitivity", isPathologyHigher, useBCACorrection, numBootstrap, cutoff)
End Function

Public Function InternalDS_CalculateSpecificity(values As Variant, pathologies As Variant, clusters As Variant, isPathologyHigher As Boolean, ParamArray extraParams() As Variant) As Double
    InternalDS_CalculateSpecificity = DS_ROC_Specificity(DS_Filter(values, pathologies, 1), DS_Filter(values, pathologies, 0), extraParams(0)(0), isPathologyHigher)
End Function

Public Function DS_ROCCorr_Specificity_Bootstrap(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional useBCACorrection As Boolean = True, Optional numBootstrap As Integer = 500) As Double()
    DS_ROCCorr_Specificity_Bootstrap = DS_ROCCorr_Bootstrap_Generic(measurementRange, pathologyRange, clusterRange, "InternalDS_CalculateSpecificity", isPathologyHigher, useBCACorrection, numBootstrap, cutoff)
End Function

Public Function InternalDS_CalculateAccuracy(values As Variant, pathologies As Variant, clusters As Variant, isPathologyHigher As Boolean, ParamArray extraParams() As Variant) As Double
    InternalDS_CalculateAccuracy = DS_ROC_Accuracy(DS_Filter(values, pathologies, 1), DS_Filter(values, pathologies, 0), extraParams(0)(0), isPathologyHigher)
End Function

Public Function DS_ROCCorr_Accuracy_Bootstrap(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional useBCACorrection As Boolean = True, Optional numBootstrap As Integer = 500) As Double()
    DS_ROCCorr_Accuracy_Bootstrap = DS_ROCCorr_Bootstrap_Generic(measurementRange, pathologyRange, clusterRange, "InternalDS_CalculateAccuracy", isPathologyHigher, useBCACorrection, numBootstrap, cutoff)
End Function

Private Function DS_ROCCorr_Bootstrap_Generic(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, calcFunc As String, ByVal isPathologyHigher As Boolean, ByVal useBCACorrection As Boolean, ByVal numBootstrap As Integer, ParamArray extraParams() As Variant) As Double() 'only 1 extra param is allowed
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim i As Long
    Dim results() As Double
    Dim bootstrap() As Variant
    Dim bootstrapSampleValues() As Variant
    Dim bootstrapSamplePathologies() As Variant
    Dim bootstrapSampleClusters() As Variant
    Dim returnValue(0 To 2) As Double
    Dim originalStat As Double
    Dim jackknifeStats() As Double
    Dim z0 As Double, a As Double
    Dim normalQuantileLower As Double, normalQuantileUpper As Double
    Dim correctedLowerPercentile As Double, correctedUpperPercentile As Double
    
    ReDim results(1 To numBootstrap)
    
    ' Make random function deterministic
    Rnd (-1)
    Randomize (123)
    
    ' Convert measurementRange to array if it's a range
    If TypeOf measurementRange Is Range Then
        arrMeasure = DS_RangeToArray(measurementRange)
    Else
        arrMeasure = measurementRange
    End If

    ' Convert pathologyRange to array if it's a range
    If TypeOf pathologyRange Is Range Then
        arrPathology = DS_RangeToArray(pathologyRange)
    Else
        arrPathology = pathologyRange
    End If

    ' Convert clusterRange to array if it's a range
    If TypeOf clusterRange Is Range Then
        arrCluster = DS_RangeToArray(clusterRange)
    Else
        arrCluster = clusterRange
    End If
    
    ' Bootstrap resampling
    For i = LBound(results) To UBound(results)
        bootstrap = DS_BootstrapClusterSample(arrMeasure, arrPathology, arrCluster)
        
        If LBound(bootstrap) = 0 And UBound(bootstrap) = 2 Then 'if this is not the case, smth went wrong
            bootstrapSampleValues = bootstrap(0)
            bootstrapSamplePathologies = bootstrap(1)
            bootstrapSampleClusters = bootstrap(2)
            
            ' Call the passed function with optional extra parameters (cutoff, etc.)
            results(i) = Application.Run(calcFunc, bootstrapSampleValues, bootstrapSamplePathologies, bootstrapSampleClusters, isPathologyHigher, extraParams)
        Else
            results(i) = 0
        End If
    Next i
    
    
    returnValue(0) = WorksheetFunction.Average(results)
    
    ' Calculate the original statistic (without resampling)
    originalStat = Application.Run(calcFunc, arrMeasure, arrPathology, arrCluster, isPathologyHigher, extraParams)
    
    ' BCa correction logic
    If useBCACorrection Then
        ' Step 1: Bias-correction factor (z0)
        Dim lessThanOriginalCount As Long
        lessThanOriginalCount = 0
        For i = LBound(results) To UBound(results)
            If results(i) < originalStat Then lessThanOriginalCount = lessThanOriginalCount + 1
        Next i
        z0 = WorksheetFunction.NormSInv(lessThanOriginalCount / numBootstrap) ' Bias correction
        
        ' Step 2: Acceleration factor (a) via Jackknife resampling on cluster level
        Dim uniqueClusterNames() As Variant
        uniqueClusterNames = DS_GetUnique(arrCluster)
        
        ReDim jackknifeStats(LBound(uniqueClusterNames) To UBound(uniqueClusterNames))
        For i = LBound(uniqueClusterNames) To UBound(uniqueClusterNames)
            Dim jackknifeSampleValues() As Variant
            Dim jackknifePathologyValues() As Variant
            Dim jackknifeClusterValues() As Variant
            
            jackknifeSampleValues = DS_JackknifeClusterSample(arrMeasure, arrCluster, uniqueClusterNames(i))
            jackknifePathologyValues = DS_JackknifeClusterSample(arrPathology, arrCluster, uniqueClusterNames(i))
            jackknifeClusterValues = DS_JackknifeClusterSample(arrCluster, arrCluster, uniqueClusterNames(i))
            jackknifeStats(i) = Application.Run(calcFunc, jackknifeSampleValues, jackknifePathologyValues, jackknifeClusterValues, isPathologyHigher, extraParams)
        Next i
        a = DS_CalculateAcceleration(jackknifeStats, originalStat)
        
        ' Step 3: Adjust percentiles for BCa correction
        normalQuantileLower = WorksheetFunction.NormSInv(0.025) ' Normal quantile for 2.5%
        normalQuantileUpper = WorksheetFunction.NormSInv(0.975) ' Normal quantile for 97.5%
        
        correctedLowerPercentile = WorksheetFunction.NormSDist(z0 + (z0 + normalQuantileLower) / (1 - a * (z0 + normalQuantileLower)))
        correctedUpperPercentile = WorksheetFunction.NormSDist(z0 + (z0 + normalQuantileUpper) / (1 - a * (z0 + normalQuantileUpper)))
        
        ' Sort the bootstrap results
        DS_QuickSort results, LBound(results), UBound(results)
        
        ' Step 4: Calculate BCa-corrected confidence intervals
        returnValue(1) = DS_Percentile(results, correctedLowerPercentile) ' Lower bound
        returnValue(2) = DS_Percentile(results, correctedUpperPercentile) ' Upper bound
    Else
        ' Standard percentile-based confidence intervals
        DS_QuickSort results, LBound(results), UBound(results)
        returnValue(1) = DS_Percentile(results, 0.025) ' Lower bound (2.5th percentile)
        returnValue(2) = DS_Percentile(results, 0.975) ' Upper bound (97.5th percentile)
    End If

    DS_ROCCorr_Bootstrap_Generic = returnValue
End Function

Public Function DS_ROCCorr_AUC_Bootstrap(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional isPathologyHigher As Boolean = True, Optional useBCACorrection As Boolean = True, Optional numBootstrap As Integer = 500) As Double()
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim i As Long
    Dim results() As Double
    Dim bootstrap() As Variant
    Dim bootstrapSampleValues() As Variant
    Dim bootstrapSamplePathologies() As Variant
    Dim bootstrapSampleClusters() As Variant
    Dim returnValue(0 To 3) As Double
    Dim originalAUC As Double
    Dim jackknifeAUCs() As Double
    Dim z0 As Double, a As Double
    Dim normalQuantileLower As Double, normalQuantileUpper As Double
    Dim correctedLowerPercentile As Double, correctedUpperPercentile As Double
    Dim extremeCount As Long ' For p-value calculation
    
    ReDim results(1 To numBootstrap)
    
    'make random function deterministic
    Rnd (-1)
    Randomize (123)
    
    ' Convert measurementRange to array if it's a range
    If TypeOf measurementRange Is Range Then
        arrMeasure = DS_RangeToArray(measurementRange)
    Else
        arrMeasure = measurementRange
    End If

    ' Convert pathologyRange to array if it's a range
    If TypeOf pathologyRange Is Range Then
        arrPathology = DS_RangeToArray(pathologyRange)
    Else
        arrPathology = pathologyRange
    End If

    ' Convert clusterRange to array if it's a range
    If TypeOf clusterRange Is Range Then
        arrCluster = DS_RangeToArray(clusterRange)
    Else
        arrCluster = clusterRange
    End If
    
    
    For i = LBound(results) To UBound(results)
        bootstrap = DS_BootstrapClusterSample(arrMeasure, arrPathology, arrCluster)
        
        If LBound(bootstrap) = 0 And UBound(bootstrap) = 2 Then 'if this is not the case, smth went wrong
            bootstrapSampleValues = bootstrap(0)
            bootstrapSamplePathologies = bootstrap(1)
            bootstrapSampleClusters = bootstrap(2)
            
            results(i) = DS_ROC_AUC(DS_Filter(bootstrapSampleValues, bootstrapSamplePathologies, 1), DS_Filter(bootstrapSampleValues, bootstrapSamplePathologies, 0), isPathologyHigher)
        
            ' Count extreme bootstrap AUC values for p-value (Nullhypothesis: no discrimantory ability -> AUC 0.5)
            If results(i) <= 0.5 Then
                extremeCount = extremeCount + 1
            End If
        Else
            results(i) = 0
        End If
    Next i
    
    
    returnValue(0) = WorksheetFunction.Average(results)
    ' Calculate p-value
    returnValue(3) = (extremeCount / numBootstrap)
    
    
    ' Calculate the original AUC (without resampling)
    originalAUC = DS_ROC_AUC(DS_Filter(arrMeasure, arrPathology, 1), DS_Filter(arrMeasure, arrPathology, 0), isPathologyHigher)
    
    ' BCa correction logic
    If useBCACorrection Then
        ' Step 1: Bias-correction factor (z0)
        Dim lessThanOriginalCount As Long
        lessThanOriginalCount = 0
        For i = LBound(results) To UBound(results)
            If results(i) < originalAUC Then lessThanOriginalCount = lessThanOriginalCount + 1
        Next i
        z0 = WorksheetFunction.NormSInv(lessThanOriginalCount / numBootstrap) ' Bias correction
        
        ' Step 2: Acceleration factor (a) via Jackknife resampling on cluster level
        Dim uniqueClusterNames() As Variant
        uniqueClusterNames = DS_GetUnique(arrCluster)
        
        ReDim jackknifeAUCs(LBound(uniqueClusterNames) To UBound(uniqueClusterNames))
        For i = LBound(uniqueClusterNames) To UBound(uniqueClusterNames)
            Dim jackknifeSampleValues() As Variant
            Dim jackknifePathologyValues() As Variant
            
            jackknifeSampleValues = DS_JackknifeClusterSample(arrMeasure, arrCluster, uniqueClusterNames(i))
            jackknifePathologyValues = DS_JackknifeClusterSample(arrPathology, arrCluster, uniqueClusterNames(i))
            jackknifeAUCs(i) = DS_ROC_AUC(DS_Filter(jackknifeSampleValues, jackknifePathologyValues, 1), DS_Filter(jackknifeSampleValues, jackknifePathologyValues, 0), isPathologyHigher)
        Next i
        a = DS_CalculateAcceleration(jackknifeAUCs, originalAUC)
        
        ' Step 3: Adjust percentiles for BCa correction
        normalQuantileLower = WorksheetFunction.NormSInv(0.025) ' Normal quantile for 2.5%
        normalQuantileUpper = WorksheetFunction.NormSInv(0.975) ' Normal quantile for 97.5%
        
        correctedLowerPercentile = WorksheetFunction.NormSDist(z0 + (z0 + normalQuantileLower) / (1 - a * (z0 + normalQuantileLower)))
        correctedUpperPercentile = WorksheetFunction.NormSDist(z0 + (z0 + normalQuantileUpper) / (1 - a * (z0 + normalQuantileUpper)))
        
        ' Sort the bootstrap results
        DS_QuickSort results, LBound(results), UBound(results)
        
        ' Step 4: Calculate BCa-corrected confidence intervals
        returnValue(1) = DS_Percentile(results, correctedLowerPercentile) ' Lower bound
        returnValue(2) = DS_Percentile(results, correctedUpperPercentile) ' Upper bound
    Else
        ' Standard percentile-based confidence intervals
        DS_QuickSort results, LBound(results), UBound(results)
        returnValue(1) = DS_Percentile(results, 0.025) ' Lower bound (2.5th percentile)
        returnValue(2) = DS_Percentile(results, 0.975) ' Upper bound (97.5th percentile)
    End If
    
    
    DS_ROCCorr_AUC_Bootstrap = returnValue
End Function
