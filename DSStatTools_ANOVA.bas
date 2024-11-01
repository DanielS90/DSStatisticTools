Attribute VB_Name = "DSStatTools_ANOVA"
Public Function DS_KruskalWallisP(ByVal valueRange1 As Range, ByVal valueRange2 As Range, Optional valueRange3 As Variant, Optional valueRange4 As Variant, Optional valueRange5 As Variant, Optional valueRange6 As Variant, Optional valueRange7 As Variant, Optional valueRange8 As Variant, Optional valueRange9 As Variant, Optional valueRange10 As Variant)
    
    Dim allValues As Variant
    allValues = DS_MergeRangesToArray(valueRange1, valueRange2, valueRange3, valueRange4, valueRange5, valueRange6, valueRange7, valueRange8, valueRange9, valueRange10)
    
    Dim categories As Integer
    categories = 0
    Dim currentRank As Double
    Dim sumR2n As Double
    sumR2n = 0
    
    Dim R1 As Double
    R1 = 0
    Dim n1 As Integer
    n1 = 0
    For Each currentCell In valueRange1
        If Not IsEmpty(currentCell) Then
            currentRank = DS_Rank(currentCell.value, allValues)
            R1 = R1 + currentRank
            n1 = n1 + 1
        End If
    Next currentCell
    sumR2n = sumR2n + (R1 * R1) / n1
    categories = categories + 1
    
    Dim R2 As Double
    R2 = 0
    Dim n2 As Integer
    n2 = 0
    For Each currentCell In valueRange2
        If Not IsEmpty(currentCell) Then
            currentRank = DS_Rank(currentCell.value, allValues)
            R2 = R2 + currentRank
            n2 = n2 + 1
        End If
    Next currentCell
    sumR2n = sumR2n + (R2 * R2) / n2
    categories = categories + 1
    
    Dim R3 As Double
    R3 = 0
    Dim n3 As Integer
    n3 = 0
    If Not IsMissing(valueRange3) Then
        For Each currentCell In valueRange3
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R3 = R3 + currentRank
                n3 = n3 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R3 * R3) / n3
        categories = categories + 1
    End If
    
    Dim R4 As Double
    R4 = 0
    Dim n4 As Integer
    n4 = 0
    If Not IsMissing(valueRange4) Then
        For Each currentCell In valueRange4
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R4 = R4 + currentRank
                n4 = n4 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R4 * R4) / n4
        categories = categories + 1
    End If
    
    Dim R5 As Double
    R5 = 0
    Dim n5 As Integer
    n5 = 0
    If Not IsMissing(valueRange5) Then
        For Each currentCell In valueRange5
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R5 = R5 + currentRank
                n5 = n5 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R5 * R5) / n5
        categories = categories + 1
    End If
    
    Dim R6 As Double
    R6 = 0
    Dim n6 As Integer
    n6 = 0
    If Not IsMissing(valueRange6) Then
        For Each currentCell In valueRange6
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R6 = R6 + currentRank
                n6 = n6 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R6 * R6) / n6
        categories = categories + 1
    End If
    
    Dim R7 As Double
    R7 = 0
    Dim n7 As Integer
    n7 = 0
    If Not IsMissing(valueRange7) Then
        For Each currentCell In valueRange7
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R7 = R7 + currentRank
                n7 = n7 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R7 * R7) / n7
        categories = categories + 1
    End If
    
    Dim R8 As Double
    R8 = 0
    Dim n8 As Integer
    n8 = 0
    If Not IsMissing(valueRange8) Then
        For Each currentCell In valueRange8
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R8 = R8 + currentRank
                n8 = n8 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R8 * R8) / n8
        categories = categories + 1
    End If
    
    Dim R9 As Double
    R9 = 0
    Dim n9 As Integer
    n9 = 0
    If Not IsMissing(valueRange9) Then
        For Each currentCell In valueRange9
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R9 = R9 + currentRank
                n9 = n9 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R9 * R9) / n9
        categories = categories + 1
    End If
    
    Dim R10 As Double
    R10 = 0
    Dim n10 As Integer
    n10 = 0
    If Not IsMissing(valueRange10) Then
        For Each currentCell In valueRange10
            If Not IsEmpty(currentCell) Then
                currentRank = DS_Rank(currentCell.value, allValues)
                R10 = R10 + currentRank
                n10 = n10 + 1
            End If
        Next currentCell
        sumR2n = sumR2n + (R10 * R10) / n10
        categories = categories + 1
    End If
    
    Dim totalN As Integer
    totalN = UBound(allValues) - LBound(allValues) + 1
    
    Dim H As Double
    H = 12 / (totalN * (totalN + 1)) * sumR2n - 3 * (totalN + 1)
    
    DS_KruskalWallisP = WorksheetFunction.ChiSq_Dist_RT(H, categories - 1)
End Function

Public Function DS_ANOVAOneWayP(ByVal valueRange1 As Range, ByVal valueRange2 As Range, Optional valueRange3 As Variant, Optional valueRange4 As Variant, Optional valueRange5 As Variant, Optional valueRange6 As Variant, Optional valueRange7 As Variant, Optional valueRange8 As Variant, Optional valueRange9 As Variant, Optional valueRange10 As Variant)
    Dim totalN As Integer
    totalN = 0
    
    Dim totalMean As Double
    totalMean = 0
    
    Dim totalGroups As Integer
    totalGroups = 0
    
    
    Dim n1 As Integer
    n1 = 0
    Dim mean1 As Double
    mean1 = 0
    For Each currentCell In valueRange1
        If Not IsEmpty(currentCell) Then
            totalMean = totalMean + currentCell.value
            totalN = totalN + 1
            mean1 = mean1 + currentCell.value
            n1 = n1 + 1
        End If
    Next currentCell
    mean1 = mean1 / n1
    totalGroups = totalGroups + 1
    
    Dim n2 As Integer
    n2 = 0
    Dim mean2 As Double
    mean2 = 0
    For Each currentCell In valueRange2
        If Not IsEmpty(currentCell) Then
            totalMean = totalMean + currentCell.value
            totalN = totalN + 1
            mean2 = mean2 + currentCell.value
            n2 = n2 + 1
        End If
    Next currentCell
    mean2 = mean2 / n2
    totalGroups = totalGroups + 1
    
    Dim n3 As Integer
    n3 = 0
    Dim mean3 As Double
    mean3 = 0
    If Not IsMissing(valueRange3) Then
        For Each currentCell In valueRange3
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean3 = mean3 + currentCell.value
                n3 = n3 + 1
            End If
        Next currentCell
        mean3 = mean3 / n3
        totalGroups = totalGroups + 1
    End If
    
    Dim n4 As Integer
    n4 = 0
    Dim mean4 As Double
    mean4 = 0
    If Not IsMissing(valueRange4) Then
        For Each currentCell In valueRange4
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean4 = mean4 + currentCell.value
                n4 = n4 + 1
            End If
        Next currentCell
        mean4 = mean4 / n4
        totalGroups = totalGroups + 1
    End If
    
    Dim n5 As Integer
    n5 = 0
    Dim mean5 As Double
    mean5 = 0
    If Not IsMissing(valueRange5) Then
        For Each currentCell In valueRange5
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean5 = mean5 + currentCell.value
                n5 = n5 + 1
            End If
        Next currentCell
        mean5 = mean5 / n5
        totalGroups = totalGroups + 1
    End If
    
    Dim n6 As Integer
    n6 = 0
    Dim mean6 As Double
    mean6 = 0
    If Not IsMissing(valueRange6) Then
        For Each currentCell In valueRange6
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean6 = mean6 + currentCell.value
                n6 = n6 + 1
            End If
        Next currentCell
        mean6 = mean6 / n6
        totalGroups = totalGroups + 1
    End If
    
    Dim n7 As Integer
    n7 = 0
    Dim mean7 As Double
    mean7 = 0
    If Not IsMissing(valueRange7) Then
        For Each currentCell In valueRange7
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean7 = mean7 + currentCell.value
                n7 = n7 + 1
            End If
        Next currentCell
        mean7 = mean7 / n7
        totalGroups = totalGroups + 1
    End If
    
    Dim n8 As Integer
    n8 = 0
    Dim mean8 As Double
    mean8 = 0
    If Not IsMissing(valueRange8) Then
        For Each currentCell In valueRange8
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean8 = mean8 + currentCell.value
                n8 = n8 + 1
            End If
        Next currentCell
        mean8 = mean8 / n8
        totalGroups = totalGroups + 1
    End If
    
    Dim n9 As Integer
    n9 = 0
    Dim mean9 As Double
    mean9 = 0
    If Not IsMissing(valueRange9) Then
        For Each currentCell In valueRange9
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean9 = mean9 + currentCell.value
                n9 = n9 + 1
            End If
        Next currentCell
        mean9 = mean9 / n9
        totalGroups = totalGroups + 1
    End If
    
    Dim n10 As Integer
    n10 = 0
    Dim mean10 As Double
    mean10 = 0
    If Not IsMissing(valueRange10) Then
        For Each currentCell In valueRange10
            If Not IsEmpty(currentCell) Then
                totalMean = totalMean + currentCell.value
                totalN = totalN + 1
                mean10 = mean10 + currentCell.value
                n10 = n10 + 1
            End If
        Next currentCell
        mean10 = mean10 / n10
        totalGroups = totalGroups + 1
    End If
    
    
    totalMean = totalMean / totalN
    
    Dim SSR As Double
    SSR = 0
    Dim SSE As Double
    SSE = 0
    
    
    SSR = SSR + n1 * (mean1 - totalMean) ^ 2
    For Each currentCell In valueRange1
        If Not IsEmpty(currentCell) Then
            SSE = SSE + (currentCell.value - mean1) ^ 2
        End If
    Next currentCell
    
    SSR = SSR + n2 * (mean2 - totalMean) ^ 2
    For Each currentCell In valueRange2
        If Not IsEmpty(currentCell) Then
            SSE = SSE + (currentCell.value - mean2) ^ 2
        End If
    Next currentCell
    
    If Not IsMissing(valueRange3) Then
        SSR = SSR + n3 * (mean3 - totalMean) ^ 2
        For Each currentCell In valueRange3
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean3) ^ 2
            End If
        Next currentCell
    End If
    
    If Not IsMissing(valueRange4) Then
        SSR = SSR + n4 * (mean4 - totalMean) ^ 2
        For Each currentCell In valueRange4
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean4) ^ 2
            End If
        Next currentCell
    End If
    
    If Not IsMissing(valueRange5) Then
        SSR = SSR + n5 * (mean5 - totalMean) ^ 2
        For Each currentCell In valueRange5
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean5) ^ 2
            End If
        Next currentCell
    End If
    
    If Not IsMissing(valueRange6) Then
        SSR = SSR + n6 * (mean6 - totalMean) ^ 2
        For Each currentCell In valueRange6
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean6) ^ 2
            End If
        Next currentCell
    End If
    
    If Not IsMissing(valueRange7) Then
        SSR = SSR + n7 * (mean7 - totalMean) ^ 2
        For Each currentCell In valueRange7
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean7) ^ 2
            End If
        Next currentCell
    End If
    
    If Not IsMissing(valueRange8) Then
        SSR = SSR + n8 * (mean8 - totalMean) ^ 2
        For Each currentCell In valueRange8
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean8) ^ 2
            End If
        Next currentCell
    End If
    
    If Not IsMissing(valueRange9) Then
        SSR = SSR + n9 * (mean9 - totalMean) ^ 2
        For Each currentCell In valueRange9
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean9) ^ 2
            End If
        Next currentCell
    End If
    
    If Not IsMissing(valueRange10) Then
        SSR = SSR + n10 * (mean10 - totalMean) ^ 2
        For Each currentCell In valueRange10
            If Not IsEmpty(currentCell) Then
                SSE = SSE + (currentCell.value - mean10) ^ 2
            End If
        Next currentCell
    End If
    
    
    Dim SST As Double
    SST = SSR + SSE
    
    Dim groupsSS As Double
    groupsSS = SSR
    
    Dim groupsDf As Double
    groupsDf = totalGroups - 1
    
    Dim groupsMS As Double
    groupsMS = groupsSS / groupsDf
    
    Dim errorSS As Double
    errorSS = SSE
    
    Dim errorDf As Double
    errorDf = totalN - totalGroups
    
    Dim errorMS As Double
    errorMS = errorSS / errorDf
    
    Dim F As Double
    F = groupsMS / errorMS
    
    DS_ANOVAOneWayP = WorksheetFunction.F_Dist_RT(F, groupsDf, errorDf)
End Function

