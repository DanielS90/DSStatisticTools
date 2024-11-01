Attribute VB_Name = "DSStatTools_ROCCorr"
Public Function DS_ROCCorr_PrintCutoffs(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional numVals As Integer = 100) As Variant
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim minVal As Double
    Dim maxVal As Double
    Dim cutoffArray() As Double
    Dim i As Integer
    Dim increment As Double
    Dim padding As Double

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

    ' Check that all three arrays have the same length
    If UBound(arrMeasure) <> UBound(arrPathology) Or UBound(arrMeasure) <> UBound(arrCluster) Then
        Err.Raise vbObjectError + 513, , "Input arrays must be of the same length."
    End If

    ' Get the minimum and maximum values from the measurement array
    minVal = WorksheetFunction.Min(arrMeasure)
    maxVal = WorksheetFunction.Max(arrMeasure)

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
    DS_ROCCorr_PrintCutoffs = cutoffArray
End Function

Public Function DS_ROCCorr_PrintTPRs(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional isPathologyHigher As Boolean = True, Optional numVals As Integer = 100) As Variant
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim cutoffArray() As Double
    Dim truePositives As Long
    Dim tprArray() As Double
    Dim i As Integer
    Dim currentCutoff As Double
    Dim totalPos As Long

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

    ' Check that all three arrays have the same length
    If UBound(arrMeasure) <> UBound(arrPathology) Or UBound(arrMeasure) <> UBound(arrCluster) Then
        Err.Raise vbObjectError + 513, , "Input arrays must be of the same length."
    End If

    ' Get the cutoff values using the DS_ROCCorr_PrintCutoffs function
    cutoffArray = DS_ROCCorr_PrintCutoffs(arrMeasure, arrPathology, arrCluster, numVals)

    ' Initialize the TPR array to hold the values
    ReDim tprArray(0 To UBound(cutoffArray))

    ' Total positives (pathology group labeled as 1)
    totalPos = DS_Occurrences(arrPathology, 1)

    ' Loop through each cutoff value to calculate TPR
    For i = LBound(cutoffArray) To UBound(cutoffArray)
        currentCutoff = cutoffArray(i)

        ' Calculate true positives based on whether higher values indicate pathology
        truePositives = 0
        Dim j As Long
        For j = LBound(arrMeasure) To UBound(arrMeasure)
            If arrPathology(j) = 1 Then
                If isPathologyHigher Then
                    If arrMeasure(j) >= currentCutoff Then
                        truePositives = truePositives + 1 ' True Positive
                    End If
                Else
                    If arrMeasure(j) <= currentCutoff Then
                        truePositives = truePositives + 1 ' True Positive
                    End If
                End If
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
    DS_ROCCorr_PrintTPRs = tprArray
End Function

Public Function DS_ROCCorr_PrintFPRs(ByVal measurementRange As Variant, ByVal pathologyRange As Variant, ByVal clusterRange As Variant, Optional isPathologyHigher As Boolean = True, Optional numVals As Integer = 100) As Variant
    Dim arrMeasure() As Variant
    Dim arrPathology() As Variant
    Dim arrCluster() As Variant
    Dim cutoffArray() As Double
    Dim falsePositives As Long
    Dim fprArray() As Double
    Dim i As Integer
    Dim currentCutoff As Double
    Dim totalNeg As Long

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

    ' Check that all three arrays have the same length
    If UBound(arrMeasure) <> UBound(arrPathology) Or UBound(arrMeasure) <> UBound(arrCluster) Then
        Err.Raise vbObjectError + 513, , "Input arrays must be of the same length."
    End If

    ' Get the cutoff values using the DS_ROCCorr_PrintCutoffs function
    cutoffArray = DS_ROCCorr_PrintCutoffs(arrMeasure, arrPathology, arrCluster, numVals)

    ' Initialize the FPR array to hold the values
    ReDim fprArray(0 To UBound(cutoffArray))

    ' Total negatives (pathology group labeled as 0)
    totalNeg = DS_Occurrences(arrPathology, 0)

    ' Loop through each cutoff value to calculate FPR
    For i = LBound(cutoffArray) To UBound(cutoffArray)
        currentCutoff = cutoffArray(i)

        ' Calculate false positives based on whether higher values indicate pathology
        falsePositives = 0
        Dim j As Long
        For j = LBound(arrMeasure) To UBound(arrMeasure)
            If arrPathology(j) = 0 Then
                If isPathologyHigher Then
                    If arrMeasure(j) >= currentCutoff Then
                        falsePositives = falsePositives + 1 ' False Positive
                    End If
                Else
                    If arrMeasure(j) <= currentCutoff Then
                        falsePositives = falsePositives + 1 ' False Positive
                    End If
                End If
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
    DS_ROCCorr_PrintFPRs = fprArray
End Function
