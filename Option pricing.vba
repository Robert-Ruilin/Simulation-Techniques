Option Explicit

Public Sub Sim_generating_quaterly()
    Dim runs As Long
    Dim tmp As Double
    Dim value As Double
    Dim i As Long, j As Long
    Dim k As Double, s0 As Double, vol As Double, r As Double, t As Double
    Dim inc As Double
    k = 100
    s0 = 100
    vol = 0.2
    r = 0.05
    t = 1
    
    runs = 10000
    Randomize
    For i = 1 To runs
        tmp = s0
        Cells(21 + i, 3) = tmp
        For j = 1 To 4
            tmp = tmp * Exp((r - 0.5 * (vol ^ 2)) * 0.25 * t + vol * Sqr(0.25 * t) * WorksheetFunction.NormSInv(Rnd()))
            Cells(21 + i, 3 + j) = tmp
        Next j
    Next i
End Sub
Public Sub exercise_payoff()
    Dim i As Long, j As Long
    Dim k As Double
    k = 100
    For i = 1 To 10000
        For j = 1 To 5
            Cells(21 + i, 2 + j) = Application.Max(k - Cells(21 + i, 2 + j), 0)
        Next j
        
    Next i

End Sub

Public Sub sort_9m()
    Range("c22:g10021").Sort key1:=Range("F22:F10021"), order1:=xlDescending, Header:=xlNo
End Sub

Public Sub value_change_9m()
    Dim stock_arr_current As Variant
    Dim stock_arr_future As Variant
    Dim val_arr(1 To 10000) As Variant
    Dim tmp_arr_current(1 To 100) As Variant
    Dim tmp_arr_future(1 To 100) As Variant
    Dim k As Double, r As Double, t As Double
    Dim tmp_compare As Double
    k = 100
    r = 0.05
    t = 1

    stock_arr_current = Application.Transpose(Range("F22:F10021"))
    stock_arr_future = Application.Transpose(Range("g22:g10021"))
    
    Dim i As Long, j As Long, m As Long
    For i = 1 To 100
    
        For j = 1 To 100
            tmp_arr_current(j) = stock_arr_current(100 * (i - 1) + j)
            tmp_arr_future(j) = stock_arr_future(100 * (i - 1) + j)
        Next j
        
        For m = 1 To 100
            tmp_compare = Exp(-r * 0.25 * t) * Application.Average(tmp_arr_future)
            If tmp_arr_current(m) < tmp_compare Then
                tmp_arr_current(m) = tmp_compare
            End If
            val_arr(100 * (i - 1) + m) = tmp_arr_current(m)
        Next m
        

    Next i
    
    Dim l As Long
    For l = 1 To 10000
        Cells(22 + l - 1, 6) = val_arr(l)
    Next l
    
End Sub

Public Sub sort_6m()
    Range("c22:g10021").Sort key1:=Range("e22:e10021"), order1:=xlDescending, Header:=xlNo
End Sub

Public Sub value_change_6m()
    Dim stock_arr_current As Variant
    Dim stock_arr_future As Variant
    Dim val_arr(1 To 10000) As Variant
    Dim tmp_arr_current(1 To 100) As Variant
    Dim tmp_arr_future(1 To 100) As Variant
    Dim k As Double, r As Double, t As Double
    Dim tmp_compare As Double
    k = 100
    r = 0.05
    t = 1

    stock_arr_current = Application.Transpose(Range("E22:E10021"))
    stock_arr_future = Application.Transpose(Range("F22:F10021"))
    
    Dim i As Long, j As Long, m As Long
    For i = 1 To 100
    
        For j = 1 To 100
            tmp_arr_current(j) = stock_arr_current(100 * (i - 1) + j)
            tmp_arr_future(j) = stock_arr_future(100 * (i - 1) + j)
        Next j
        
        For m = 1 To 100
            tmp_compare = Exp(-r * 0.25 * t) * Application.Average(tmp_arr_future)
            If tmp_arr_current(m) < tmp_compare Then
                tmp_arr_current(m) = tmp_compare
            End If
            val_arr(100 * (i - 1) + m) = tmp_arr_current(m)
        Next m
        

    Next i
    
    Dim l As Long
    For l = 1 To 10000
        Cells(22 + l - 1, 5) = val_arr(l)
    Next l
    
End Sub
Public Sub sort_3m()
    Range("c22:g10021").Sort key1:=Range("D22:D10021"), order1:=xlDescending, Header:=xlNo
End Sub

Public Sub value_change_3m()
    Dim stock_arr_current As Variant
    Dim stock_arr_future As Variant
    Dim val_arr(1 To 10000) As Variant
    Dim tmp_arr_current(1 To 100) As Variant
    Dim tmp_arr_future(1 To 100) As Variant
    Dim k As Double, r As Double, t As Double
    Dim tmp_compare As Double
    k = 100
    r = 0.05
    t = 1

    stock_arr_current = Application.Transpose(Range("D22:D10021"))
    stock_arr_future = Application.Transpose(Range("E22:E10021"))
    
    Dim i As Long, j As Long, m As Long
    For i = 1 To 100
    
        For j = 1 To 100
            tmp_arr_current(j) = stock_arr_current(100 * (i - 1) + j)
            tmp_arr_future(j) = stock_arr_future(100 * (i - 1) + j)
        Next j
        
        For m = 1 To 100
            tmp_compare = Exp(-r * 0.25 * t) * Application.Average(tmp_arr_future)
            If tmp_arr_current(m) < tmp_compare Then
                tmp_arr_current(m) = tmp_compare
            End If
            val_arr(100 * (i - 1) + m) = tmp_arr_current(m)
        Next m
        

    Next i
    
    Dim l As Long
    For l = 1 To 10000
        Cells(22 + l - 1, 4) = val_arr(l)
    Next l
    
End Sub
Sub american_put()
    Dim american_put As Double
    Dim i As Long
    Dim runs As Long
    Dim r As Double, t As Double
    Dim temp As Double
    runs = 10000
    r = 0.05
    t = 1
    For i = 1 To runs
        american_put = american_put + Exp(-r * 0.25 * t) * Cells(22 + i - 1, 4)
    Next i
    american_put = american_put / runs
    Cells(11, 1) = american_put
End Sub
