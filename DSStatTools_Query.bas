Option Compare Text

Public Function DS_Select(ByVal cellRange As Range, ByVal conditionRange As Range, ByVal comparison As String)
    If Not cellRange.Rows.Count = conditionRange.Rows.Count Then
        Exit Function
    End If
    
    Dim result() As Variant
    
    Dim iR As Integer
    Dim counter As Integer
    counter = 0
    For iR = 1 To cellRange.Rows.Count
        If DS_PatternMatch(conditionRange(iR, 1), comparison) Then
            ReDim Preserve result(0 To counter)
            result(counter) = cellRange(iR, 1)
            counter = counter + 1
        End If
    Next iR
    
    DS_Select = result
End Function

Public Function DS_SelectAND(ByVal cellRange As Range, ByVal conditionRange As Range, ByVal comp As String, Optional comp2 As Variant, Optional comp3 As Variant, Optional comp4 As Variant)
    If Not cellRange.Rows.Count = conditionRange.Rows.Count Then
        Exit Function
    End If
    
    Dim result() As Variant
    
    Dim iR As Integer
    Dim counter As Integer
    counter = 0
    For iR = 1 To cellRange.Rows.Count
        Dim match As Boolean
        match = False
        
        If DS_PatternMatch(conditionRange(iR, 1), comp) Then
            match = True
        End If
        
        If Not IsMissing(comp2) Then
            If match And DS_PatternMatch(conditionRange(iR, 1), comp2) Then
                match = True
            Else
                match = False
            End If
        End If
        
        If Not IsMissing(comp3) Then
            If match And DS_PatternMatch(conditionRange(iR, 1), comp3) Then
                match = True
            Else
                match = False
            End If
        End If
        
        If Not IsMissing(comp4) Then
            If match And DS_PatternMatch(conditionRange(iR, 1), comp4) Then
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
    
    DS_SelectAND = result
End Function

Public Function DS_SelectOR(ByVal cellRange As Range, ByVal conditionRange As Range, ByVal comp As String, Optional comp2 As Variant, Optional comp3 As Variant, Optional comp4 As Variant)
    If Not cellRange.Rows.Count = conditionRange.Rows.Count Then
        Exit Function
    End If
    
    Dim result() As Variant
    
    Dim iR As Integer
    Dim counter As Integer
    counter = 0
    For iR = 1 To cellRange.Rows.Count
        Dim match As Boolean
        match = False
        
        If DS_PatternMatch(conditionRange(iR, 1), comp) Then
            match = True
        End If
        
        If Not IsMissing(comp2) Then
            If DS_PatternMatch(conditionRange(iR, 1), comp2) Then
                match = True
            End If
        End If
        
        If Not IsMissing(comp3) Then
            If DS_PatternMatch(conditionRange(iR, 1), comp3) Then
                match = True
            End If
        End If
        
        If Not IsMissing(comp4) Then
            If DS_PatternMatch(conditionRange(iR, 1), comp4) Then
                match = True
            End If
        End If
        
        If match Then
            ReDim Preserve result(0 To counter)
            result(counter) = cellRange(iR, 1)
            counter = counter + 1
        End If
    Next iR
    
    DS_SelectOR = result
End Function

