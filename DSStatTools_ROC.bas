Attribute VB_Name = "DSStatTools_ROC"
Public Function DS_ROC_PrintCutoffs(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional numVals As Integer = 100) As Variant
    Dim arr1() As Variant
    Dim arr2() As Variant
    Dim minVal As Double
    Dim maxVal As Double
    Dim cutoffArray() As Double
    Dim i As Integer
    Dim increment As Double
    Dim padding As Double
    
    ' Convert cellRange1 to array if it is a range
    If TypeOf cellRange1 Is Range Then
        arr1 = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arr1 = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range
    If TypeOf cellRange2 Is Range Then
        arr2 = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arr2 = cellRange2 ' Already an array
    End If
    
    ' Get the minimum and maximum values from both arrays
    minVal = WorksheetFunction.Min(Application.Min(arr1), Application.Min(arr2))
    maxVal = WorksheetFunction.Max(Application.Max(arr1), Application.Max(arr2))
    
    ' Define a small padding value (e.g., 1% of the range)
    padding = (maxVal - minVal) * 0.01 ' 1% padding
    
    ' Adjust the min and max values with padding
    minVal = minVal - padding
    maxVal = maxVal + padding
    
    ' Calculate the increment based on numVals
    increment = (maxVal - minVal) / (numVals - 1)
    
    ' Initialize the cutoff array to hold the values
    ReDim cutoffArray(0 To numVals - 1)
    
    ' Fill the cutoff array with values from minVal to maxVal
    For i = 0 To numVals - 1
        cutoffArray(i) = minVal + i * increment
    Next i
    
    ' Return the cutoff array as the function result
    DS_ROC_PrintCutoffs = cutoffArray
End Function

Public Function DS_ROC_PrintCutoffsWithoutOutliers(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional numVals As Integer = 100) As Variant
    Dim arr1() As Variant
    Dim arr2() As Variant
    Dim minVal As Double
    Dim maxVal As Double
    Dim cutoffArray() As Double
    Dim i As Integer
    Dim increment As Double
    Dim padding As Double
    
    ' Convert cellRange1 to array if it is a range
    If TypeOf cellRange1 Is Range Then
        arr1 = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arr1 = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range
    If TypeOf cellRange2 Is Range Then
        arr2 = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arr2 = cellRange2 ' Already an array
    End If
    
    ' Get the minimum and maximum values from both arrays
    minVal = WorksheetFunction.Min(Application.Min(arr1), Application.Min(arr2))
    maxVal = WorksheetFunction.Max(Application.Max(arr1), Application.Max(arr2))
    
    ' Calculate the increment based on numVals
    increment = (maxVal - minVal) / (numVals - 1)
    
    ' Initialize the cutoff array to hold the values
    ReDim cutoffArray(0 To numVals - 1)
    
    ' Fill the cutoff array with values from minVal to maxVal
    For i = 0 To numVals - 1
        cutoffArray(i) = minVal + i * increment
    Next i
    
    ' Return the cutoff array as the function result
    DS_ROC_PrintCutoffsWithoutOutliers = cutoffArray
End Function

Public Function DS_ROC_PrintTPRs(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True, Optional numVals As Integer = 100) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim cutoffArray() As Double
    Dim truePositives As Long
    Dim tprArray() As Double
    Dim i As Integer
    Dim currentCutoff As Double
    Dim totalPos As Long
    
    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If
    
    ' Get the cutoff values using the DS_ROC_PrintCutoffs function
    cutoffArray = DS_ROC_PrintCutoffs(arrPos, arrNeg, numVals)
    
    ' Initialize the TPR array to hold the values
    ReDim tprArray(0 To UBound(cutoffArray))
    
    ' Total positives
    totalPos = UBound(arrPos) - LBound(arrPos) + 1
    
    ' Loop through each cutoff value to calculate TPR
    For i = LBound(cutoffArray) To UBound(cutoffArray)
        currentCutoff = cutoffArray(i)
        
        ' Calculate true positives
        truePositives = 0
        Dim j As Long
        For j = LBound(arrPos) To UBound(arrPos)
            If (isPathologyHigher And arrPos(j) >= currentCutoff) Or _
               (Not isPathologyHigher And arrPos(j) <= currentCutoff) Then
                truePositives = truePositives + 1 ' True Positive
            End If
        Next j
        
        ' Calculate TPR
        If totalPos > 0 Then
            tprArray(i) = truePositives / totalPos ' True Positive Rate
        Else
            tprArray(i) = 0 ' Avoid division by zero
        End If
    Next i
    
    ' Return the TPR array as the function result
    DS_ROC_PrintTPRs = tprArray
End Function

Public Function DS_ROC_PrintFPRs(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True, Optional numVals As Integer = 100) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim cutoffArray() As Double
    Dim falsePositives As Long
    Dim fprArray() As Double
    Dim i As Integer
    Dim currentCutoff As Double
    Dim totalNeg As Long
    
    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If
    
    ' Get the cutoff values using the DS_ROC_PrintCutoffs function
    cutoffArray = DS_ROC_PrintCutoffs(arrPos, arrNeg, numVals)
    
    ' Initialize the FPR array to hold the values
    ReDim fprArray(0 To UBound(cutoffArray))
    
    ' Total negatives
    totalNeg = UBound(arrNeg) - LBound(arrNeg) + 1
    
    ' Loop through each cutoff value to calculate FPR
    For i = LBound(cutoffArray) To UBound(cutoffArray)
        currentCutoff = cutoffArray(i)
        
        ' Calculate false positives
        falsePositives = 0
        
        Dim j As Long
        For j = LBound(arrNeg) To UBound(arrNeg)
            If (isPathologyHigher And arrNeg(j) >= currentCutoff) Or _
               (Not isPathologyHigher And arrNeg(j) <= currentCutoff) Then
                falsePositives = falsePositives + 1 ' False Positive
            End If
        Next j
        
        ' Calculate FPR
        If totalNeg > 0 Then
            fprArray(i) = falsePositives / totalNeg ' False Positive Rate
        Else
            fprArray(i) = 0 ' Avoid division by zero
        End If
    Next i
    
    ' Return the FPR array as the function result
    DS_ROC_PrintFPRs = fprArray
End Function

Public Function DS_ROC_AUC_Trapezoidal(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True, Optional numVals As Integer = 1000) As Double
    Dim tprArray() As Double
    Dim fprArray() As Double
    Dim AUC As Double
    Dim i As Integer
    Dim totalPoints As Integer
    
    'simple trapezoidal implementation. It's generally preferable to use the DeLong method provided below

    ' Get TPRs and FPRs
    tprArray = DS_ROC_PrintTPRs(cellRange1, cellRange2, isPathologyHigher, numVals)
    fprArray = DS_ROC_PrintFPRs(cellRange1, cellRange2, isPathologyHigher, numVals)

    ' Initialize AUC
    AUC = 0
    totalPoints = UBound(tprArray) - LBound(tprArray) + 1

    ' Calculate AUC using the trapezoidal rule
    For i = LBound(tprArray) To UBound(tprArray) - 1
        AUC = AUC + (fprArray(i + 1) - fprArray(i)) * (tprArray(i + 1) + tprArray(i)) / 2
    Next i

    ' Return the AUC value
    DS_ROC_AUC_Trapezoidal = Math.Abs(AUC)
End Function

Public Function DS_ROC_AUC(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True) As Double
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim combinedScores() As Variant
    Dim combinedLabels() As Integer
    Dim ranks() As Double
    Dim nPos As Long
    Dim nNeg As Long
    Dim i As Long, j As Long
    Dim sumRanksPos As Double
    Dim AUC As Double

    ' Convert cellRange1 and cellRange2 to arrays
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1)
    Else
        arrPos = cellRange1
    End If

    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2)
    Else
        arrNeg = cellRange2
    End If

    ' Get the number of positives and negatives
    nPos = UBound(arrPos) - LBound(arrPos) + 1
    nNeg = UBound(arrNeg) - LBound(arrNeg) + 1

    ' Combine scores and labels
    ReDim combinedScores(1 To nPos + nNeg)
    ReDim combinedLabels(1 To nPos + nNeg)

    ' Add positive samples
    For i = 1 To nPos
        combinedScores(i) = arrPos(LBound(arrPos) + i - 1)
        combinedLabels(i) = 1 ' Positive label
    Next i

    ' Add negative samples
    For i = 1 To nNeg
        combinedScores(nPos + i) = arrNeg(LBound(arrNeg) + i - 1)
        combinedLabels(nPos + i) = 0 ' Negative label
    Next i

    ' Adjust scores if isPathologyHigher is False
    If Not isPathologyHigher Then
        For i = 1 To UBound(combinedScores)
            combinedScores(i) = -combinedScores(i)
        Next i
    End If

    ' Assign ranks to combined scores
    ranks = DS_ROC_Helpers_GetRanks(combinedScores)

    ' Sum ranks for positive samples
    sumRanksPos = 0
    For i = 1 To UBound(combinedScores)
        If combinedLabels(i) = 1 Then
            sumRanksPos = sumRanksPos + ranks(i)
        End If
    Next i

    ' Calculate AUC using Mann-Whitney U statistic
    AUC = (sumRanksPos - nPos * (nPos + 1) / 2) / (nPos * nNeg)

    ' Return the AUC value
    DS_ROC_AUC = AUC
End Function

Public Function DS_ROC_AUCStdErr(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True) As Double
    Dim AUC As Double
    Dim aucVar As Double
    Dim se As Double

    ' Calculate AUC
    AUC = DS_ROC_AUC(cellRange1, cellRange2, isPathologyHigher)

    ' Calculate variance of AUC using DeLong's method
    aucVar = DS_ROC_AUCDeLongVar(cellRange1, cellRange2, AUC, isPathologyHigher)

    ' Calculate Standard Error
    If aucVar > 0 Then
        se = Sqr(aucVar)
    Else
        se = 0
    End If

    ' Return the Standard Error
    DS_ROC_AUCStdErr = se
End Function

Public Function DS_ROC_AUCCI(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True, Optional significanceLevel As Double = 0.95) As Variant
    Dim AUC As Double
    Dim se As Double
    Dim zValue As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    
    'uses the DeLong method

    ' Calculate AUC
    AUC = DS_ROC_AUC(cellRange1, cellRange2, isPathologyHigher)
    
    ' Calculate Standard Error
    se = DS_ROC_AUCStdErr(cellRange1, cellRange2, isPathologyHigher)

    ' Calculate Z-value for the given significance level
    zValue = Application.WorksheetFunction.NormSInv(1 - (1 - significanceLevel) / 2)

    ' Calculate confidence interval bounds
    lowerBound = AUC - zValue * se
    upperBound = AUC + zValue * se

    ' Ensure bounds are within the interval [0, 1]
    If lowerBound < 0 Then lowerBound = 0
    If upperBound > 1 Then upperBound = 1

    ' Return the confidence interval as an array
    Dim confidenceInterval(0 To 1) As Double
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_AUCCI = confidenceInterval
End Function
Public Function DS_ROC_AUCDeLongP(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True) As Double
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim nPos As Long
    Dim nNeg As Long
    Dim AUC As Double
    Dim aucVar As Double
    Dim Z As Double
    Dim pValue As Double
    
    'calculates the p value for an AUC using DeLong's method
    
    ' Convert cellRange1 (positives) and cellRange2 (negatives) to arrays
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1)
    Else
        arrPos = cellRange1
    End If
    
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2)
    Else
        arrNeg = cellRange2
    End If
    
    ' Get the number of positives and negatives
    nPos = UBound(arrPos) - LBound(arrPos) + 1
    nNeg = UBound(arrNeg) - LBound(arrNeg) + 1

    ' Step 1: Calculate AUC
    AUC = DS_ROC_AUC(arrPos, arrNeg, isPathologyHigher)

    ' Step 2: Calculate the variance of AUC using DeLong's method
    aucVar = DS_ROC_AUCDeLongVar(arrPos, arrNeg, AUC, isPathologyHigher)

    ' Step 3: Calculate z-score for the AUC (comparing it to 0.5)
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

    ' Return the p-value
    DS_ROC_AUCDeLongP = pValue
End Function

Public Function DS_ROC_AUCDeLongVar(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal AUC As Double, Optional isPathologyHigher As Boolean = True) As Double
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim nPos As Long
    Dim nNeg As Long
    Dim vPos() As Double
    Dim vNeg() As Double
    Dim i As Long, j As Long
    Dim V_p As Double
    Dim V_n As Double
    Dim aucVar As Double
    Dim sumPos As Double
    Dim sumNeg As Double

    ' Convert cellRange1 (positives) and cellRange2 (negatives) to arrays
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1)
    Else
        arrPos = cellRange1
    End If

    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2)
    Else
        arrNeg = cellRange2
    End If

    ' Get the number of positives and negatives
    nPos = UBound(arrPos) - LBound(arrPos) + 1
    nNeg = UBound(arrNeg) - LBound(arrNeg) + 1

    ' Ensure we have more than one positive and negative sample
    If nPos <= 1 Or nNeg <= 1 Then
        DS_ROC_AUCDeLongVar = 0
        Exit Function
    End If

    ' Initialize the variance components
    ReDim vPos(LBound(arrPos) To UBound(arrPos))
    ReDim vNeg(LBound(arrNeg) To UBound(arrNeg))

    ' Step 1: Calculate the rank-based comparisons for positive cases
    For i = LBound(arrPos) To UBound(arrPos)
        sumPos = 0
        For j = LBound(arrNeg) To UBound(arrNeg)
            If (isPathologyHigher And arrPos(i) >= arrNeg(j)) Or (Not isPathologyHigher And arrPos(i) <= arrNeg(j)) Then
                sumPos = sumPos + 1
            End If
        Next j
        vPos(i) = sumPos / nNeg
    Next i

    ' Step 2: Calculate the rank-based comparisons for negative cases
    For j = LBound(arrNeg) To UBound(arrNeg)
        sumNeg = 0
        For i = LBound(arrPos) To UBound(arrPos)
            If (isPathologyHigher And arrNeg(j) < arrPos(i)) Or (Not isPathologyHigher And arrNeg(j) > arrPos(i)) Then
                sumNeg = sumNeg + 1
            End If
        Next i
        vNeg(j) = sumNeg / nPos
    Next j

    ' Step 3: Compute V_p and V_n
    V_p = 0
    For i = LBound(arrPos) To UBound(arrPos)
        V_p = V_p + (vPos(i) - AUC) ^ 2
    Next i
    V_p = V_p / (nPos - 1)

    V_n = 0
    For j = LBound(arrNeg) To UBound(arrNeg)
        V_n = V_n + (vNeg(j) - AUC) ^ 2
    Next j
    V_n = V_n / (nNeg - 1)

    ' Step 4: Compute the variance of AUC based on DeLong's method
    aucVar = V_p / nPos + V_n / nNeg

    ' Return the calculated variance
    DS_ROC_AUCDeLongVar = aucVar
End Function

Public Function DS_ROC_Sensitivity(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True) As Double
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim truePositives As Long
    Dim totalPos As Long
    Dim sensitivity As Double
    Dim i As Long

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Total positives
    totalPos = UBound(arrPos) - LBound(arrPos) + 1

    ' Calculate true positives
    truePositives = 0
    For i = LBound(arrPos) To UBound(arrPos)
        If (isPathologyHigher And arrPos(i) >= cutoff) Or _
           (Not isPathologyHigher And arrPos(i) <= cutoff) Then
            truePositives = truePositives + 1 ' True Positive
        End If
    Next i

    ' Calculate sensitivity (True Positive Rate)
    If totalPos > 0 Then
        sensitivity = truePositives / totalPos
    Else
        sensitivity = 0 ' Avoid division by zero
    End If

    ' Return the sensitivity
    DS_ROC_Sensitivity = sensitivity
End Function

Public Function DS_ROC_Specificity(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True) As Double
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim trueNegatives As Long
    Dim totalNeg As Long
    Dim specificity As Double
    Dim i As Long

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Total negatives
    totalNeg = UBound(arrNeg) - LBound(arrNeg) + 1

    ' Calculate true negatives
    trueNegatives = 0
    For i = LBound(arrNeg) To UBound(arrNeg)
        If (isPathologyHigher And arrNeg(i) < cutoff) Or _
           (Not isPathologyHigher And arrNeg(i) > cutoff) Then
            trueNegatives = trueNegatives + 1 ' True Negative
        End If
    Next i

    ' Calculate specificity (True Negative Rate)
    If totalNeg > 0 Then
        specificity = trueNegatives / totalNeg
    Else
        specificity = 0 ' Avoid division by zero
    End If

    ' Return the specificity
    DS_ROC_Specificity = specificity
End Function

Public Function DS_ROC_Accuracy(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True) As Double
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim truePositives As Long
    Dim trueNegatives As Long
    Dim totalCases As Long
    Dim accuracy As Double
    Dim i As Long

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Calculate true positives and true negatives
    truePositives = 0
    trueNegatives = 0
    For i = LBound(arrPos) To UBound(arrPos)
        If (isPathologyHigher And arrPos(i) >= cutoff) Or _
           (Not isPathologyHigher And arrPos(i) <= cutoff) Then
            truePositives = truePositives + 1 ' True Positive
        End If
    Next i
    
    For i = LBound(arrNeg) To UBound(arrNeg)
        If (isPathologyHigher And arrNeg(i) < cutoff) Or _
           (Not isPathologyHigher And arrNeg(i) > cutoff) Then
            trueNegatives = trueNegatives + 1 ' True Negative
        End If
    Next i

    ' Total cases
    totalCases = (UBound(arrPos) - LBound(arrPos) + 1) + (UBound(arrNeg) - LBound(arrNeg) + 1)

    ' Calculate accuracy
    If totalCases > 0 Then
        accuracy = (truePositives + trueNegatives) / totalCases
    Else
        accuracy = 0 ' Avoid division by zero
    End If

    ' Return accuracy
    DS_ROC_Accuracy = accuracy
End Function

Public Function DS_ROC_SensitivityCI_Exact(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional confidenceLevel As Double = 0.95) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim truePositives As Long
    Dim totalPos As Long
    Dim alpha As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim sensitivity As Double
    Dim confidenceInterval(0 To 1) As Double
    Dim i As Long
    
    'calculates the CI using the exact binomial confidence interval (Clopper-Pearson method). This is generally suitable for lower sample sizes (< 30-50) and when the true sensitivity might be at the extremes

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Total positives
    totalPos = UBound(arrPos) - LBound(arrPos) + 1

    ' Calculate true positives
    truePositives = 0
    For i = LBound(arrPos) To UBound(arrPos)
        If (isPathologyHigher And arrPos(i) >= cutoff) Or _
           (Not isPathologyHigher And arrPos(i) <= cutoff) Then
            truePositives = truePositives + 1 ' True Positive
        End If
    Next i

    ' Calculate sensitivity
    If totalPos > 0 Then
        sensitivity = truePositives / totalPos
    Else
        sensitivity = 0 ' Avoid division by zero
    End If

    ' Alpha value for the confidence level
    alpha = 1 - confidenceLevel

    ' Calculate lower bound using Clopper-Pearson exact method
    If truePositives = 0 Then
        lowerBound = 0 ' No true positives, lower bound is 0
    Else
        lowerBound = Application.WorksheetFunction.BetaInv(alpha / 2, truePositives, totalPos - truePositives + 1)
    End If

    ' Calculate upper bound using Clopper-Pearson exact method
    If truePositives = totalPos Then
        upperBound = 1 ' If all positives are true positives, upper bound is 1
    Else
        upperBound = Application.WorksheetFunction.BetaInv(1 - alpha / 2, truePositives + 1, totalPos - truePositives)
    End If

    ' Return confidence interval as an array
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_SensitivityCI_Exact = confidenceInterval
End Function

Public Function DS_ROC_SensitivityCI(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional confidenceLevel As Double = 0.95) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim truePositives As Long
    Dim totalPos As Long
    Dim sensitivity As Double
    Dim alpha As Double
    Dim Z As Double
    Dim center As Double
    Dim marginOfError As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim confidenceInterval(0 To 1) As Double
    Dim i As Long
    
    'calculates the CI using the Wilson Score method. This is generally advisable for larger smaple sizes (> 30) and when the expected true value is near 50% as it provides a narrower estimate.

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Total positives
    totalPos = UBound(arrPos) - LBound(arrPos) + 1

    ' Calculate true positives
    truePositives = 0
    For i = LBound(arrPos) To UBound(arrPos)
        If (isPathologyHigher And arrPos(i) >= cutoff) Or _
           (Not isPathologyHigher And arrPos(i) <= cutoff) Then
            truePositives = truePositives + 1 ' True Positive
        End If
    Next i

    ' Calculate sensitivity
    If totalPos > 0 Then
        sensitivity = truePositives / totalPos
    Else
        sensitivity = 0 ' Avoid division by zero
    End If

    ' Calculate z value for the confidence level (e.g., 1.96 for 95% confidence)
    alpha = 1 - confidenceLevel
    Z = Application.WorksheetFunction.NormSInv(1 - alpha / 2)

    ' Calculate the Wilson score confidence interval
    center = (sensitivity + (Z ^ 2) / (2 * totalPos)) / (1 + (Z ^ 2) / totalPos)
    marginOfError = Z * Sqr((sensitivity * (1 - sensitivity) / totalPos) + (Z ^ 2) / (4 * totalPos ^ 2)) / (1 + (Z ^ 2) / totalPos)

    ' Calculate lower and upper bounds
    lowerBound = center - marginOfError
    upperBound = center + marginOfError

    ' Ensure bounds are within the interval [0, 1]
    If lowerBound < 0 Then lowerBound = 0
    If upperBound > 1 Then upperBound = 1

    ' Return confidence interval as an array
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_SensitivityCI = confidenceInterval
End Function

Public Function DS_ROC_SpecificityCI_Exact(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional confidenceLevel As Double = 0.95) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim trueNegatives As Long
    Dim totalNeg As Long
    Dim alpha As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim specificity As Double
    Dim confidenceInterval(0 To 1) As Double
    Dim i As Long
    
    'calculates the CI using the exact binomial confidence interval (Clopper-Pearson method)

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Total negatives
    totalNeg = UBound(arrNeg) - LBound(arrNeg) + 1

    ' Calculate true negatives
    trueNegatives = 0
    For i = LBound(arrNeg) To UBound(arrNeg)
        If (isPathologyHigher And arrNeg(i) < cutoff) Or _
           (Not isPathologyHigher And arrNeg(i) > cutoff) Then
            trueNegatives = trueNegatives + 1 ' True Negative
        End If
    Next i

    ' Calculate specificity
    If totalNeg > 0 Then
        specificity = trueNegatives / totalNeg
    Else
        specificity = 0 ' Avoid division by zero
    End If

    ' Alpha value for the confidence level
    alpha = 1 - confidenceLevel

    ' Calculate lower bound using Clopper-Pearson exact method
    If trueNegatives = 0 Then
        lowerBound = 0 ' No true negatives, lower bound is 0
    Else
        lowerBound = Application.WorksheetFunction.BetaInv(alpha / 2, trueNegatives, totalNeg - trueNegatives + 1)
    End If

    ' Calculate upper bound using Clopper-Pearson exact method
    If trueNegatives = totalNeg Then
        upperBound = 1 ' If all negatives are true negatives, upper bound is 1
    Else
        upperBound = Application.WorksheetFunction.BetaInv(1 - alpha / 2, trueNegatives + 1, totalNeg - trueNegatives)
    End If

    ' Return confidence interval as an array
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_SpecificityCI_Exact = confidenceInterval
End Function

Public Function DS_ROC_SpecificityCI(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional confidenceLevel As Double = 0.95) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim trueNegatives As Long
    Dim totalNeg As Long
    Dim specificity As Double
    Dim alpha As Double
    Dim Z As Double
    Dim center As Double
    Dim marginOfError As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim confidenceInterval(0 To 1) As Double
    Dim i As Long
    
    'calculates the CI using the Wilson Score method

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Total negatives
    totalNeg = UBound(arrNeg) - LBound(arrNeg) + 1

    ' Calculate true negatives
    trueNegatives = 0
    For i = LBound(arrNeg) To UBound(arrNeg)
        If (isPathologyHigher And arrNeg(i) < cutoff) Or _
           (Not isPathologyHigher And arrNeg(i) > cutoff) Then
            trueNegatives = trueNegatives + 1 ' True Negative
        End If
    Next i

    ' Calculate specificity
    If totalNeg > 0 Then
        specificity = trueNegatives / totalNeg
    Else
        specificity = 0 ' Avoid division by zero
    End If

    ' Calculate z value for the confidence level (e.g., 1.96 for 95% confidence)
    alpha = 1 - confidenceLevel
    Z = Application.WorksheetFunction.NormSInv(1 - alpha / 2)

    ' Calculate the Wilson score confidence interval
    center = (specificity + (Z ^ 2) / (2 * totalNeg)) / (1 + (Z ^ 2) / totalNeg)
    marginOfError = Z * Sqr((specificity * (1 - specificity) / totalNeg) + (Z ^ 2) / (4 * totalNeg ^ 2)) / (1 + (Z ^ 2) / totalNeg)

    ' Calculate lower and upper bounds
    lowerBound = center - marginOfError
    upperBound = center + marginOfError

    ' Ensure bounds are within the interval [0, 1]
    If lowerBound < 0 Then lowerBound = 0
    If upperBound > 1 Then upperBound = 1

    ' Return confidence interval as an array
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_SpecificityCI = confidenceInterval
End Function

Public Function DS_ROC_AccuracyCI(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional confidenceLevel As Double = 0.95) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim truePositives As Long
    Dim trueNegatives As Long
    Dim totalCases As Long
    Dim accuracy As Double
    Dim alpha As Double
    Dim Z As Double
    Dim center As Double
    Dim marginOfError As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim confidenceInterval(0 To 1) As Double
    Dim i As Long

    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Calculate true positives and true negatives
    truePositives = 0
    trueNegatives = 0
    For i = LBound(arrPos) To UBound(arrPos)
        If (isPathologyHigher And arrPos(i) >= cutoff) Or _
           (Not isPathologyHigher And arrPos(i) <= cutoff) Then
            truePositives = truePositives + 1 ' True Positive
        End If
    Next i
    
    For i = LBound(arrNeg) To UBound(arrNeg)
        If (isPathologyHigher And arrNeg(i) < cutoff) Or _
           (Not isPathologyHigher And arrNeg(i) > cutoff) Then
            trueNegatives = trueNegatives + 1 ' True Negative
        End If
    Next i

    ' Total cases
    totalCases = (UBound(arrPos) - LBound(arrPos) + 1) + (UBound(arrNeg) - LBound(arrNeg) + 1)

    ' Calculate accuracy
    If totalCases > 0 Then
        accuracy = (truePositives + trueNegatives) / totalCases
    Else
        accuracy = 0 ' Avoid division by zero
    End If

    ' Calculate z value for the confidence level (e.g., 1.96 for 95% confidence)
    alpha = 1 - confidenceLevel
    Z = Application.WorksheetFunction.NormSInv(1 - alpha / 2)

    ' Calculate the Wilson score confidence interval
    center = (accuracy + (Z ^ 2) / (2 * totalCases)) / (1 + (Z ^ 2) / totalCases)
    marginOfError = Z * Sqr((accuracy * (1 - accuracy) / totalCases) + (Z ^ 2) / (4 * totalCases ^ 2)) / (1 + (Z ^ 2) / totalCases)

    ' Calculate lower and upper bounds
    lowerBound = center - marginOfError
    upperBound = center + marginOfError

    ' Ensure bounds are within the interval [0, 1]
    If lowerBound < 0 Then lowerBound = 0
    If upperBound > 1 Then upperBound = 1

    ' Return confidence interval as an array
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_AccuracyCI = confidenceInterval
End Function

Public Function DS_ROC_AccuracyCI_Exact(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, ByVal cutoff As Double, Optional isPathologyHigher As Boolean = True, Optional confidenceLevel As Double = 0.95) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim truePositives As Long
    Dim trueNegatives As Long
    Dim totalCases As Long
    Dim accuracy As Double
    Dim alpha As Double
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim confidenceInterval(0 To 1) As Double
    Dim i As Long
    
    ' Convert cellRange1 to array if it is a range (assumed to be positives)
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1) ' Converts range to array
    Else
        arrPos = cellRange1 ' Already an array
    End If
    
    ' Convert cellRange2 to array if it is a range (assumed to be negatives)
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2) ' Converts range to array
    Else
        arrNeg = cellRange2 ' Already an array
    End If

    ' Calculate true positives and true negatives
    truePositives = 0
    trueNegatives = 0
    For i = LBound(arrPos) To UBound(arrPos)
        If (isPathologyHigher And arrPos(i) >= cutoff) Or _
           (Not isPathologyHigher And arrPos(i) <= cutoff) Then
            truePositives = truePositives + 1 ' True Positive
        End If
    Next i
    
    For i = LBound(arrNeg) To UBound(arrNeg)
        If (isPathologyHigher And arrNeg(i) < cutoff) Or _
           (Not isPathologyHigher And arrNeg(i) > cutoff) Then
            trueNegatives = trueNegatives + 1 ' True Negative
        End If
    Next i

    ' Total cases
    totalCases = (UBound(arrPos) - LBound(arrPos) + 1) + (UBound(arrNeg) - LBound(arrNeg) + 1)

    ' Calculate accuracy
    If totalCases > 0 Then
        accuracy = (truePositives + trueNegatives) / totalCases
    Else
        accuracy = 0 ' Avoid division by zero
    End If

    ' Alpha value for the confidence level
    alpha = 1 - confidenceLevel

    ' Calculate lower bound using Clopper-Pearson exact method
    If truePositives + trueNegatives = 0 Then
        lowerBound = 0 ' No correct cases, lower bound is 0
    Else
        lowerBound = Application.WorksheetFunction.BetaInv(alpha / 2, truePositives + trueNegatives, totalCases - (truePositives + trueNegatives) + 1)
    End If

    ' Calculate upper bound using Clopper-Pearson exact method
    If truePositives + trueNegatives = totalCases Then
        upperBound = 1 ' If all cases are correct, upper bound is 1
    Else
        upperBound = Application.WorksheetFunction.BetaInv(1 - alpha / 2, truePositives + trueNegatives + 1, totalCases - (truePositives + trueNegatives))
    End If

    ' Return confidence interval as an array
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_AccuracyCI_Exact = confidenceInterval
End Function

Public Function DS_ROC_YoudenCutoffCI(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True, Optional numVals As Integer = 500, Optional numBootstrap As Integer = 500, Optional confidenceLevel As Double = 0.95) As Variant
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim bootstrapCutoffs() As Double
    Dim i As Integer
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim confidenceInterval(0 To 1) As Double

    ' Convert ranges to arrays
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1)
    Else
        arrPos = cellRange1
    End If
    
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2)
    Else
        arrNeg = cellRange2
    End If
    
    'make random function deterministic
    Rnd (-1)
    Randomize (123)
    
    ' Initialize the bootstrap results array
    ReDim bootstrapCutoffs(1 To numBootstrap)
    
    ' Perform bootstrapping to get distribution of cutoffs
    For i = 1 To numBootstrap
        ' Resample with replacement to create new datasets
        Dim resampledPos() As Variant
        Dim resampledNeg() As Variant
        resampledPos = DS_BootstrapSample(arrPos)
        resampledNeg = DS_BootstrapSample(arrNeg)
        
        ' Calculate the ideal cutoff for this bootstrap sample using Youden's Index
        bootstrapCutoffs(i) = DS_ROC_OptimalCutoffYouden(resampledPos, resampledNeg, isPathologyHigher, numVals)
    Next i
    
    ' Sort the bootstrapped cutoffs
    Call DS_QuickSort(bootstrapCutoffs, LBound(bootstrapCutoffs), UBound(bootstrapCutoffs))
    
    ' Calculate the bounds of the confidence interval
    lowerBound = bootstrapCutoffs(Application.WorksheetFunction.RoundUp((1 - confidenceLevel) / 2 * numBootstrap, 0))
    upperBound = bootstrapCutoffs(Application.WorksheetFunction.RoundDown((1 - (1 - confidenceLevel) / 2) * numBootstrap, 0))
    
    ' Return the confidence interval as an array
    confidenceInterval(0) = lowerBound
    confidenceInterval(1) = upperBound

    DS_ROC_YoudenCutoffCI = confidenceInterval
End Function


' Helper function to calculate Youden's Index cutoff
Public Function DS_ROC_OptimalCutoffYouden(ByVal cellRange1 As Variant, ByVal cellRange2 As Variant, Optional isPathologyHigher As Boolean = True, Optional numVals As Integer = 1000) As Double
    Dim arrPos() As Variant
    Dim arrNeg() As Variant
    Dim cutoffArray() As Double
    Dim sensitivityArray() As Double
    Dim fprArray() As Double
    Dim youdenArray() As Double
    Dim i As Integer
    Dim maxYouden As Double
    Dim idealCutoff As Double
    Dim sumCutoff As Double
    Dim countCutoffs As Integer
    Dim currentLumpStart As Integer
    Dim currentLumpEnd As Integer
    Dim bestLumpStart As Integer
    Dim bestLumpEnd As Integer
    Dim bestLumpSize As Integer
    Const epsilon As Double = 0.000000001 '1E-9
    
    ' Convert cellRange1 (positives) and cellRange2 (negatives) to arrays
    If TypeOf cellRange1 Is Range Then
        arrPos = DS_RangeToArray(cellRange1)
    Else
        arrPos = cellRange1
    End If
    
    If TypeOf cellRange2 Is Range Then
        arrNeg = DS_RangeToArray(cellRange2)
    Else
        arrNeg = cellRange2
    End If
    
    ' Get the cutoffs and corresponding sensitivity and specificity values
    cutoffArray = DS_ROC_PrintCutoffs(arrPos, arrNeg, numVals)
    sensitivityArray = DS_ROC_PrintTPRs(arrPos, arrNeg, isPathologyHigher, numVals)
    fprArray = DS_ROC_PrintFPRs(arrPos, arrNeg, isPathologyHigher, numVals)
    
    ' Initialize the Youden index array
    ReDim youdenArray(0 To UBound(cutoffArray))
    
    ' Calculate the Youden index for each cutoff
    maxYouden = -1 ' Start with an impossibly low value
    For i = LBound(cutoffArray) To UBound(cutoffArray)
        youdenArray(i) = sensitivityArray(i) + (1 - fprArray(i)) - 1
        ' Find the maximum Youden index
        If youdenArray(i) > maxYouden Then
            maxYouden = youdenArray(i)
        End If
    Next i
    
    ' Now, find the largest contiguous group of cutoffs with the maximum Youden index
    currentLumpStart = -1
    currentLumpEnd = -1
    bestLumpStart = -1
    bestLumpEnd = -1
    bestLumpSize = 0
    countCutoffs = 0
    
    For i = LBound(cutoffArray) To UBound(cutoffArray)
        If Abs(youdenArray(i) - maxYouden) < epsilon Then
            If currentLumpStart = -1 Then
                ' Start a new lump
                currentLumpStart = i
                currentLumpEnd = i
            Else
                ' Extend the current lump
                currentLumpEnd = i
            End If
        Else
            ' End the current lump and check if it's the largest
            If currentLumpStart <> -1 Then
                If (currentLumpEnd - currentLumpStart + 1) > bestLumpSize Then
                    bestLumpStart = currentLumpStart
                    bestLumpEnd = currentLumpEnd
                    bestLumpSize = currentLumpEnd - currentLumpStart + 1
                End If
                ' Reset the current lump
                currentLumpStart = -1
                currentLumpEnd = -1
            End If
        End If
    Next i
    
    ' Handle the case where the largest lump ends at the last cutoff
    If currentLumpStart <> -1 Then
        If (currentLumpEnd - currentLumpStart + 1) > bestLumpSize Then
            bestLumpStart = currentLumpStart
            bestLumpEnd = currentLumpEnd
            bestLumpSize = currentLumpEnd - currentLumpStart + 1
        End If
    End If
    
    ' Calculate the average of the largest contiguous lump
    If bestLumpStart <> -1 Then
        sumCutoff = 0
        countCutoffs = 0
        For i = bestLumpStart To bestLumpEnd
            sumCutoff = sumCutoff + cutoffArray(i)
            countCutoffs = countCutoffs + 1
        Next i
        idealCutoff = sumCutoff / countCutoffs
    Else
        idealCutoff = 0 ' Fallback, though this case should not occur
    End If
    
    ' Return the averaged cutoff that maximizes the Youden index
    DS_ROC_OptimalCutoffYouden = idealCutoff
End Function
