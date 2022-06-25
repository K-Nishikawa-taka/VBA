Attribute VB_Name = "Module18"
Sub Next•¶()

    Dim sum As Integer
    Dim i As Integer
    
    sum = 0
    
    For i = 1 To 10
        sum = sum + i
    Next i
    
    Range("A1").Value = sum
    
End Sub
