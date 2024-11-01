Attribute VB_Name = "DSStatTools_ROC_Helpers"
Public Function DS_ROC_Helpers_GetRanks(ByRef values() As Variant) As Double()
    Dim n As Long
    Dim i As Long, j As Long
    Dim sumRanks As Double
    Dim count As Long
    Dim adjustedValues() As Variant
    Dim sortedIndices() As Long
    Dim ranks() As Double
    Dim finalRanks() As Double
    Dim lb As Long
    Dim ub As Long

    lb = LBound(values)
    ub = UBound(values)
    n = ub - lb + 1

    ReDim adjustedValues(1 To n)
    ReDim sortedIndices(1 To n)
    ReDim ranks(1 To n)

    ' Copy values to adjustedValues(1 To n)
    For i = 1 To n
        adjustedValues(i) = values(lb + i - 1)
        sortedIndices(i) = i
    Next i

    ' Sort adjustedValues and sortedIndices
    DS_ROC_Helpers_QuickSort adjustedValues, sortedIndices, 1, n

    ' Assign ranks, handling ties
    i = 1
    Do While i <= n
        count = 1
        sumRanks = i
        ' Check for ties
        Do While i < n
            If adjustedValues(sortedIndices(i)) = adjustedValues(sortedIndices(i + 1)) Then
                i = i + 1
                count = count + 1
                sumRanks = sumRanks + i
            Else
                Exit Do
            End If
        Loop
        sumRanks = sumRanks / count
        ' Assign the average rank to all tied values
        For j = i - count + 1 To i
            ranks(sortedIndices(j)) = sumRanks
        Next j
        i = i + 1
    Loop

    ' Map ranks back to the original indices
    ReDim finalRanks(lb To ub)
    For i = 1 To n
        finalRanks(lb + i - 1) = ranks(i)
    Next i

    DS_ROC_Helpers_GetRanks = finalRanks
End Function

' QuickSort algorithm to sort values and indices
Private Sub DS_ROC_Helpers_QuickSort(ByRef values() As Variant, ByRef indices() As Long, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long
    Dim pivot As Variant
    Dim tempIndex As Long

    low = first
    high = last
    pivot = values(indices((first + last) \ 2))

    Do While low <= high
        Do While values(indices(low)) < pivot
            low = low + 1
        Loop
        Do While values(indices(high)) > pivot
            high = high - 1
        Loop
        If low <= high Then
            ' Swap indices
            tempIndex = indices(low)
            indices(low) = indices(high)
            indices(high) = tempIndex
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then DS_ROC_Helpers_QuickSort values, indices, first, high
    If low < last Then DS_ROC_Helpers_QuickSort values, indices, low, last
End Sub

Public Function DS_GetUnique(ByVal arr As Variant) As Variant
    Dim uniqueItems() As Variant
    Dim i As Long
    Dim j As Long
    Dim currentValue As Variant
    Dim found As Boolean
    Dim uniqueCount As Long
    
    ' Initialize uniqueItems as an empty array
    uniqueCount = 0
    ReDim uniqueItems(0 To 0)

    ' Loop through the array to find unique values
    For i = LBound(arr) To UBound(arr)
        currentValue = arr(i)
        found = False
        
        ' Check if currentValue is already in uniqueItems
        For j = LBound(uniqueItems) To UBound(uniqueItems)
            If uniqueItems(j) = currentValue Then
                found = True
                Exit For
            End If
        Next j
        
        ' If the value is not found, add it to the uniqueItems array
        If Not found Then
            ' Resize the array to accommodate the new unique value
            If uniqueCount = 0 Then
                ReDim uniqueItems(0 To 0)
            Else
                ReDim Preserve uniqueItems(0 To uniqueCount)
            End If
            
            ' Add the currentValue to the uniqueItems array
            uniqueItems(uniqueCount) = currentValue
            uniqueCount = uniqueCount + 1
        End If
    Next i

    ' Return the array of unique items
    DS_GetUnique = uniqueItems
End Function

Function DS_Percentile(ByRef values As Variant, ByVal percentile As Double) As Double
    Dim n As Long
    Dim rank As Double
    Dim lowerIndex As Long
    Dim upperIndex As Long
    Dim fractionalPart As Double
    Dim result As Double

    ' Step 1: Sort the array using DS_Quicksort
    Call DS_QuickSort(values, LBound(values), UBound(values))

    ' Step 2: Calculate the rank of the desired percentile
    n = UBound(values) - LBound(values) + 1 ' Number of elements in the array
    rank = (n - 1) * percentile + 1 ' The rank formula for a percentile

    ' Step 3: Determine the lower and upper indices
    lowerIndex = WorksheetFunction.Floor(rank, 1)
    upperIndex = WorksheetFunction.Ceiling(rank, 1)

    ' Step 4: Interpolation between values if needed
    If lowerIndex = upperIndex Then
        ' If rank is an exact integer, return the value at that index
        result = values(lowerIndex)
    Else
        ' If rank is not an integer, interpolate between the two closest values
        fractionalPart = rank - lowerIndex
        result = values(lowerIndex) + fractionalPart * (values(upperIndex) - values(lowerIndex))
    End If

    ' Step 5: Return the interpolated value as the percentile result
    DS_Percentile = result
End Function
