Attribute VB_Name = "Module24"

Sub Subプロシージャ()

    Dim sum As Integer
    Dim i As Integer
    
    sum = 0
    
    For i = 1 To 10
        sum = sum + i
    Next i
    
    Range("A1").Value = sum
    
    Call otherCellSet

End Sub

Sub Subプロシージャ2()

    Dim multiply As Integer
    Dim i As Integer
    
    multyply = 1
    
    For i = 1 To 5
        multyply = multyply * i
    Next i
    
    Range("A1").Value = multyply
    
    Call otherCellSet

End Sub

Sub otherCellSet()
    Range("A2").Value = Range("A1") * 2
    Range("A3").Value = Range("A2") * 2
End Sub
