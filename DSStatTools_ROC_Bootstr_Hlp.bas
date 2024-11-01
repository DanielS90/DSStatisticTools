Attribute VB_Name = "DSStatTools_ROC_Bootstr_Hlp"
' Helper function to generate a bootstrap sample with replacement
Public Function DS_BootstrapSample(ByRef arr() As Variant) As Variant
    Dim i As Long
    Dim n As Long
    Dim resampledArr() As Variant
    Dim index As Long
    
    ' Get the size of the original array
    n = UBound(arr) - LBound(arr) + 1
    ReDim resampledArr(LBound(arr) To UBound(arr))
    
    ' Sample with replacement
    For i = LBound(arr) To UBound(arr)
        index = LBound(arr) + Int(Rnd() * n)
        resampledArr(i) = arr(index)
    Next i
    
    ' Return the resampled array
    DS_BootstrapSample = resampledArr
End Function

' Helper function to generate a clustered bootstrap sample with replacement returns a 2d result array. result(0) is an array containing all values, result(1) contains the associated pathology info and result(2) the cluster for each value
Public Function DS_BootstrapClusterSample(ByRef values() As Variant, ByRef pathologies() As Variant, ByRef clusters() As Variant)
    Dim i As Long
    Dim tmp As Long
    Dim j As Long
    Dim sampleClusterIndex As Long
    Dim sampleClusterValues() As Variant
    Dim sampleClusterPathologies() As Variant
    Dim nClusters As Long
    Dim uniqueClusters() As Variant
    Dim resultValues() As Variant
    Dim resultPathologies() As Variant
    Dim resultClusters() As Variant
    Dim result() As Variant
    
    If UBound(values) <> UBound(clusters) Then
        Exit Function
    End If
    
    uniqueClusters = DS_GetUnique(clusters)
    
    ' Get the size of the original array
    nClusters = UBound(uniqueClusters) - LBound(uniqueClusters) + 1
    
    ' Sample with replacement
    For i = LBound(uniqueClusters) To UBound(uniqueClusters)
        sampleClusterIndex = LBound(uniqueClusters) + Int(Rnd() * nClusters)

        sampleClusterValues = DS_Filter(values, clusters, uniqueClusters(sampleClusterIndex))
        sampleClusterPathologies = DS_Filter(pathologies, clusters, uniqueClusters(sampleClusterIndex))
        
        resultValues = DS_JoinArrays(resultValues, sampleClusterValues)
        resultPathologies = DS_JoinArrays(resultPathologies, sampleClusterPathologies)
        
        If DS_ArrayInitialized(resultClusters) Then
            tmp = UBound(resultClusters)
        Else
            tmp = LBound(resultValues)
        End If
        
        ReDim Preserve resultClusters(LBound(resultValues) To UBound(resultValues))
        
        For j = tmp To UBound(resultValues)
            resultClusters(j) = uniqueClusters(sampleClusterIndex)
        Next j
    Next i
    
    'this should never happen, but you never know
    If UBound(resultValues) <> UBound(resultPathologies) Or UBound(resultValues) <> UBound(resultClusters) Then
        Exit Function
    End If
    
    ' Return the resampled array
    ReDim result(0 To 2)
    result(0) = resultValues
    result(1) = resultPathologies
    result(2) = resultClusters
    DS_BootstrapClusterSample = result
End Function
Function DS_JackknifeSample(ByRef dataArray() As Variant, ByVal omitIndex As Long) As Variant()
    Dim n As Long
    Dim resultArray() As Variant
    Dim i As Long, j As Long
    
    n = UBound(dataArray) - LBound(dataArray) + 1 ' Length of the original array
    ReDim resultArray(1 To n - 1) ' The jackknife sample will have n-1 elements
    
    j = 1
    For i = LBound(dataArray) To UBound(dataArray)
        If i <> omitIndex Then
            resultArray(j) = dataArray(i)
            j = j + 1
        End If
    Next i
    
    DS_JackknifeSample = resultArray
End Function

Function DS_JackknifeClusterSample(ByRef values As Variant, ByRef clusters As Variant, ByVal omitClusterValue As Long) As Variant()
    Dim n As Long
    Dim resultArray() As Variant
    Dim i As Long, j As Long
    
    If UBound(values) <> UBound(clusters) Then
        Exit Function
    End If
    
    DS_JackknifeClusterSample = DS_FilterExclude(values, clusters, omitClusterValue)
End Function

Function DS_Filter(sourceArray As Variant, filterArray As Variant, filterValue As Variant)
    Dim filtered() As Variant
    Dim count As Long
    Dim i As Long
    
    If LBound(sourceArray) <> LBound(filterArray) Or UBound(sourceArray) <> UBound(filterArray) Then
        Exit Function
    End If

    count = 0
    For i = LBound(sourceArray) To UBound(sourceArray)
        If filterArray(i) = filterValue Then
            count = count + 1
            ReDim Preserve filtered(1 To count)
            filtered(count) = sourceArray(i)
        End If
    Next i

    DS_Filter = filtered
End Function

Function DS_FilterExclude(sourceArray As Variant, filterArray As Variant, excludeValue As Variant)
    Dim filtered() As Variant
    Dim count As Long
    Dim i As Long
    
    If LBound(sourceArray) <> LBound(filterArray) Or UBound(sourceArray) <> UBound(filterArray) Then
        Exit Function
    End If

    count = 0
    For i = LBound(sourceArray) To UBound(sourceArray)
        If Not filterArray(i) = excludeValue Then
            count = count + 1
            ReDim Preserve filtered(1 To count)
            filtered(count) = sourceArray(i)
        End If
    Next i

    DS_FilterExclude = filtered
End Function

Public Function DS_CalculateAcceleration(ByRef jackknifeAUCs() As Double, ByVal originalAUC As Double) As Double
    Dim meanJackknife As Double
    Dim sumCubed As Double, sumSquared As Double
    Dim i As Long
    Dim n As Long
    
    n = UBound(jackknifeAUCs) - LBound(jackknifeAUCs) + 1
    meanJackknife = WorksheetFunction.Average(jackknifeAUCs)
    
    For i = LBound(jackknifeAUCs) To UBound(jackknifeAUCs)
        sumSquared = sumSquared + (meanJackknife - jackknifeAUCs(i)) ^ 2
        sumCubed = sumCubed + (meanJackknife - jackknifeAUCs(i)) ^ 3
    Next i
    
    DS_CalculateAcceleration = sumCubed / (6 * (sumSquared ^ 1.5))
End Function
