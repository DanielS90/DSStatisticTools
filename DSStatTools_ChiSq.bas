Public Function DS_ChiSquare(ByVal cellRange As Range)
    
    Dim rowSums() As Double
    ReDim rowSums(1 To cellRange.Rows.Count)
    
    Dim colSums() As Double
    ReDim colSums(1 To cellRange.Columns.Count)
    
    Dim totalSum As Double
    
    Dim i As Integer
    For i = 1 To cellRange.Rows.Count
        Dim rowSum As Integer
        rowSum = WorksheetFunction.Sum(cellRange.Rows(i))
        rowSums(i) = rowSum
        totalSum = totalSum + rowSum
    Next i
    
    For i = 1 To cellRange.Columns.Count
        Dim colSum As Integer
        colSum = WorksheetFunction.Sum(cellRange.Columns(i))
        colSums(i) = colSum
    Next i
    
    Dim expectedValues() As Double
    ReDim expectedValues(1 To cellRange.Rows.Count, 1 To cellRange.Columns.Count)
    
    Dim rowIndex As Integer
    Dim colIndex As Integer
    
    For rowIndex = 1 To cellRange.Rows.Count
        For colIndex = 1 To cellRange.Columns.Count
            expectedValues(rowIndex, colIndex) = rowSums(rowIndex) * colSums(colIndex) / totalSum
        Next colIndex
    Next rowIndex
    
    Dim testStatisticValues() As Double
    ReDim testStatisticValues(1 To cellRange.Rows.Count, 1 To cellRange.Columns.Count)
    Dim testTotal As Double
    
    For rowIndex = 1 To cellRange.Rows.Count
        For colIndex = 1 To cellRange.Columns.Count
            testStatisticValues(rowIndex, colIndex) = (cellRange(rowIndex, colIndex) - expectedValues(rowIndex, colIndex)) ^ 2 / expectedValues(rowIndex, colIndex)
            testTotal = testTotal + testStatisticValues(rowIndex, colIndex)
        Next colIndex
    Next rowIndex
    
    DS_ChiSquare = testTotal
End Function

Public Function DS_ChiSquareDof(ByVal cellRange As Range)
    Dim dof As Integer
    dof = (cellRange.Rows.Count - 1) * (cellRange.Columns.Count - 1)
    DS_ChiSquareDof = dof
End Function

Public Function DS_ChiSquareP(ByVal cellRange As Range)
    
    Dim rowSums() As Double
    ReDim rowSums(1 To cellRange.Rows.Count)
    
    Dim colSums() As Double
    ReDim colSums(1 To cellRange.Columns.Count)
    
    Dim totalSum As Double
    
    Dim i As Integer
    For i = 1 To cellRange.Rows.Count
        Dim rowSum As Integer
        rowSum = WorksheetFunction.Sum(cellRange.Rows(i))
        rowSums(i) = rowSum
        totalSum = totalSum + rowSum
    Next i
    
    For i = 1 To cellRange.Columns.Count
        Dim colSum As Integer
        colSum = WorksheetFunction.Sum(cellRange.Columns(i))
        colSums(i) = colSum
    Next i
    
    Dim expectedValues() As Double
    ReDim expectedValues(1 To cellRange.Rows.Count, 1 To cellRange.Columns.Count)
    
    Dim rowIndex As Integer
    Dim colIndex As Integer
    
    For rowIndex = 1 To cellRange.Rows.Count
        For colIndex = 1 To cellRange.Columns.Count
            expectedValues(rowIndex, colIndex) = rowSums(rowIndex) * colSums(colIndex) / totalSum
        Next colIndex
    Next rowIndex
    
    DS_ChiSquareP = WorksheetFunction.ChiSq_Test(cellRange, expectedValues)
End Function
