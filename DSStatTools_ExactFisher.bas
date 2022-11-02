Public Function DS_ExactFisher2x2P(ByVal cellRange As Range)
    
    If Not cellRange.Rows.count = 2 Then
        Exit Function
    End If
    
    If Not cellRange.Columns.count = 2 Then
        Exit Function
    End If
    
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim p As Double
    Dim originalP As Double
    Dim counter As Integer
    Dim ua As Double
    Dim ub As Double
    Dim uc As Double
    Dim ud As Double
    Dim up As Double
    
    a = cellRange(1, 1)
    b = cellRange(1, 2)
    c = cellRange(2, 1)
    d = cellRange(2, 2)
    
    p = DS_ExactFisherPFor(a, b, c, d)
    
    originalP = p
    counter = 1
    
    ua = a
    ub = b
    uc = c
    ud = d
    While ua > 0 And ud > 0
        ua = ua - 1
        ub = ub + 1
        uc = uc + 1
        ud = ud - 1
       
        up = DS_ExactFisherPFor(ua, ub, uc, ud)
        If up <= originalP Then
            p = p + up
            counter = counter + 1
        End If
    Wend
    
    
    ua = a
    ub = b
    uc = c
    ud = d
    While ub > 0 And uc > 0
        ua = ua + 1
        ub = ub - 1
        uc = uc - 1
        ud = ud + 1
        
        up = DS_ExactFisherPFor(ua, ub, uc, ud)
        If up <= originalP Then
            p = p + up
            counter = counter + 1
        End If
    Wend
    
    DS_ExactFisher2x2P = p
    
End Function

Private Function DS_bin(ByVal n As Double, ByVal k As Double)
    Dim result As Double
    result = 1
    
    Dim i As Integer
    For i = 1 To k
        result = result * (n - (k - i))
        result = result / i
    Next i
    
    DS_bin = result
End Function

Private Function DS_ExactFisherPFor(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double)
    Dim n As Double
    n = a + b + c + d
    
    Dim x As Double
    Dim y As Double
    
    x = DS_bin(a + b, a) * DS_bin(c + d, c)
    y = DS_bin(n, a + c)
    
    Dim p As Double
    p = x / y
    
    DS_ExactFisherPFor = p
End Function
