Option Explicit


Public a As Double
Public c As Double
Public m As Double
Public xnl As Double

Function my_LCG() As Double
    xnl = (a * xnl + c) Mod m
    my_LCG = xnl / m 'divide so that in between 0 and 1
End Function

Sub init_LCG()
    a = Cells(2, 2)
    c = Cells(3, 2)
    m = Cells(4, 2)
    xnl = Cells(5, 2)
End Sub

Sub Button1_Click()
    Dim runs
    Dim i As Long
    Dim R As Double
    Call init_LCG
    runs = 10000
    
    For i = 1 To runs
        R = my_LCG()
        Cells(2 + i, 4) = R
    Next i
    
End Sub

Sub Button2_Click()
    Dim rnd_arr As Variant
    Dim frequency() As Double
    Dim i As Long, j As Long
    Dim k As Long, l As Long
    Dim temp As Double
    rnd_arr = Application.Transpose(Range("D3:D10002"))
    ReDim frequency(1 To (UBound(rnd_arr) / 100))
    
    For i = 1 To UBound(rnd_arr)
        temp = rnd_arr(i)
        For j = 1 To UBound(frequency)
            If 0.01 * (j - 1) <= temp And temp < 0.01 * j Then
                frequency(j) = frequency(j) + 1
            End If
        Next j
    Next i
    
    For k = 1 To UBound(frequency)
        Cells(2 + k, 8) = frequency(k)
    Next k
    
    For l = 1 To UBound(frequency)
        If (100 - 1.96 * Sqr(99)) < frequency(l) And frequency(l) < (100 + 1.96 * Sqr(99)) Then
            Cells(2 + l, 9) = "Accept"
        Else
            Cells(2 + l, 9) = "Reject"
        End If
    Next l

End Sub

Sub Button3_Click()
    Dim frq_arr As Variant
    Dim frq_dist() As Double
    Dim i As Long, j As Long
    Dim k As Long
    Dim temp As Double
    frq_arr = Application.Transpose(Range("H3:H102"))
    ReDim frq_dist(1 To 60)
    
    For i = 1 To UBound(frq_arr)
        temp = frq_arr(i)
        For j = 1 To UBound(frq_dist)
            If 70 + j - 1 <= temp And temp < 70 + j Then
                frq_dist(j) = frq_dist(j) + 1
            End If
        Next j
    Next i
    
    For k = 1 To UBound(frq_dist)
        Cells(2 + k, 14) = frq_dist(k)
    Next k

End Sub
Sub Button4_Click()
    Dim rnd_arr As Variant
    Dim count_condition As Long
    Dim count As Long
    Dim i As Long
    rnd_arr = Application.Transpose(Range("D3:D10002"))

    For i = 2 To UBound(rnd_arr)
        If rnd_arr(i - 1) > 0.5 Then
            count_condition = count_condition + 1
            If rnd_arr(i) > 0.5 Then
                count = count + 1
            End If
        End If
    Next i
    Cells(53, 2) = count_condition
    Cells(54, 2) = count
    Cells(55, 2) = count_condition - count
    
    Dim mu As Double
    Dim sig As Double
    mu = count_condition * 0.5
    sig = Sqr(count_condition * 0.5 * 0.5)
    
    If mu - 1.96 * sig < count And count < mu + 1.96 * sig Then
        Cells(58, 2) = "Accept"
    Else
        Cells(58, 2) = "Reject"
    End If

End Sub
Sub Button5_Click()
    Dim rnd_arr As Variant
    Dim i As Long
    Dim count As Long
    Dim count_inc As Long
    Dim count_dec As Long
    rnd_arr = Application.Transpose(Range("D3:D10002"))

    For i = 4 To UBound(rnd_arr)
        If rnd_arr(i - 2) > rnd_arr(i - 3) And rnd_arr(i - 1) > rnd_arr(i - 2) Then
            count = count + 1
            If rnd_arr(i) > rnd_arr(i - 1) Then
                count_inc = count_inc + 1
            Else
                count_dec = count_dec + 1
            End If
        End If
    Next i
    Cells(65, 2) = count
    Cells(66, 2) = count_dec
    Cells(67, 2) = count_inc
    
    Dim mu As Double
    Dim sig As Double
    mu = count * 0.75
    sig = Sqr(count * 0.75 * 0.25)
    
    If mu - 1.96 * sig < count_dec And count_dec < mu + 1.96 * sig Then
        Cells(70, 2) = "Accept"
    Else
        Cells(70, 2) = "Reject"
    End If

End Sub
