Attribute VB_Name = "Module22"

Sub 多次元配列()

    Dim tensuu(3, 2) As Integer
    
    tensuu(0, 0) = 78
    tensuu(0, 1) = 45
    tensuu(0, 2) = 52
    
    tensuu(1, 0) = 98
    tensuu(1, 1) = 87
    tensuu(1, 2) = 84
    
    tensuu(2, 0) = 52
    tensuu(2, 1) = 45
    tensuu(2, 2) = 87
    
    tensuu(3, 0) = 71
    tensuu(3, 1) = 92
    tensuu(3, 2) = 90
    
    Dim sum As Integer
    Dim heikin As Integer
    Dim i As Integer
    Dim j As Integer
    
    sum = 0
    
    For i = 0 To 3
        For j = 0 To 2
            sum = sum + tensuu(i, j)
        Next j
    Next i
    
    heikin = sum / 12
    
    Range("A1").Value = "全生徒の全教科の平均点は" & heikin

End Sub
