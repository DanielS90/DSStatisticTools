Attribute VB_Name = "DSStatTools_Query"
Option Compare Text

Public Function DS_Select(ByVal cellRange As Range, ByVal conditionRange As Range, ByVal comparison As String)
    If Not cellRange.Rows.count = conditionRange.Rows.count Then
        Exit Function
    End If
    
    Dim result() As Variant
    
    Dim iR As Integer
    Dim counter As Integer
    counter = 0
    For iR = 1 To cellRange.Rows.count
        If DS_PatternMatch(conditionRange(iR, 1), comparison) Then
            ReDim Preserve result(0 To counter)
            result(counter) = cellRange(iR, 1)
            counter = counter + 1
        End If
    Next iR
    
    If counter > 0 Then
        DS_Select = result
    Else
        DS_Select = Empty
    End If
End Function

Public Function DS_SelectAND(ByVal cellRange As Range, ByVal conditionRange As Range, ByVal comp As String, Optional conditionRange2 As Variant, Optional comp2 As Variant, Optional conditionRange3 As Variant, Optional comp3 As Variant, Optional conditionRange4 As Variant, Optional comp4 As Variant)
    If Not cellRange.Rows.count = conditionRange.Rows.count Then
        Exit Function
    End If
    
    If Not IsMissing(conditionRange2) Then
        If Not conditionRange2.Rows.count = conditionRange.Rows.count Then
            Exit Function
        End If
    End If
    
    If Not IsMissing(conditionRange3) Then
        If Not conditionRange3.Rows.count = conditionRange.Rows.count Then
            Exit Function
        End If
    End If
    
    If Not IsMissing(conditionRange4) Then
        If Not conditionRange4.Rows.count = conditionRange.Rows.count Then
            Exit Function
        End If
    End If
    
    Dim result() As Variant
    
    Dim iR As Integer
    Dim counter As Integer
    counter = 0
    For iR = 1 To cellRange.Rows.count
        Dim match As Boolean
        match = False
        
        If DS_PatternMatch(conditionRange(iR, 1), comp) Then
            match = True
        End If
        
        If Not IsMissing(comp2) Then
            If match And DS_PatternMatch(conditionRange2(iR, 1), comp2) Then
                match = True
            Else
                match = False
            End If
        End If
        
        If Not IsMissing(comp3) Then
            If match And DS_PatternMatch(conditionRange3(iR, 1), comp3) Then
                match = True
            Else
                match = False
            End If
        End If
        
        If Not IsMissing(comp4) Then
            If match And DS_PatternMatch(conditionRange4(iR, 1), comp4) Then
                match = True
            Else
                match = False
            End If
        End If
        
        If match Then
            ReDim Preserve result(0 To counter)
            result(counter) = cellRange(iR, 1)
            counter = counter + 1
        End If
    Next iR
    
    If counter > 0 Then
        DS_SelectAND = result
    Else
        DS_SelectAND = Empty
    End If
End Function

Public Function DS_SelectOR(ByVal cellRange As Range, ByVal conditionRange As Range, ByVal comp As String, Optional conditionRange2 As Variant, Optional comp2 As Variant, Optional conditionRange3 As Variant, Optional comp3 As Variant, Optional conditionRange4 As Variant, Optional comp4 As Variant)
    If Not cellRange.Rows.count = conditionRange.Rows.count Then
        Exit Function
    End If
    
    If Not IsMissing(conditionRange2) Then
        If Not conditionRange2.Rows.count = conditionRange.Rows.count Then
            Exit Function
        End If
    End If
    
    If Not IsMissing(conditionRange3) Then
        If Not conditionRange3.Rows.count = conditionRange.Rows.count Then
            Exit Function
        End If
    End If
    
    If Not IsMissing(conditionRange4) Then
        If Not conditionRange4.Rows.count = conditionRange.Rows.count Then
            Exit Function
        End If
    End If
    
    Dim result() As Variant
    
    Dim iR As Integer
    Dim counter As Integer
    counter = 0
    For iR = 1 To cellRange.Rows.count
        Dim match As Boolean
        match = False
        
        If DS_PatternMatch(conditionRange(iR, 1), comp) Then
            match = True
        End If
        
        If Not IsMissing(comp2) Then
            If DS_PatternMatch(conditionRange2(iR, 1), comp2) Then
                match = True
            End If
        End If
        
        If Not IsMissing(comp3) Then
            If DS_PatternMatch(conditionRange3(iR, 1), comp3) Then
                match = True
            End If
        End If
        
        If Not IsMissing(comp4) Then
            If DS_PatternMatch(conditionRange4(iR, 1), comp4) Then
                match = True
            End If
        End If
        
        If match Then
            ReDim Preserve result(0 To counter)
            result(counter) = cellRange(iR, 1)
            counter = counter + 1
        End If
    Next iR
    
    If counter > 0 Then
        DS_SelectOR = result
    Else
        DS_SelectOR = Empty
    End If
End Function

Public Function DS_UniqueValues(ByVal cellRange As Variant)
    If TypeOf cellRange Is Range Then
        cellRange = DS_RangeToArray(cellRange)
    End If
    
    Dim result() As Variant
    ReDim result(0)
    
    Dim valueCounter As Integer
    valueCounter = 0
    
    Dim val As Variant
    For Each val In cellRange
        If Not DS_ValueInArray(val, result) Then
            ReDim Preserve result(valueCounter)
            result(UBound(result)) = val
            valueCounter = valueCounter + 1
        End If
    Next val
    
    If valueCounter > 0 Then
        DS_UniqueValues = result
    Else
        DS_UniqueValues = Empty
    End If
End Function

