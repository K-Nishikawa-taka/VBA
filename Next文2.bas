Attribute VB_Name = "Module19"

Sub Next•¶2()

    Dim sum As Integer
    Dim i As Integer
    
    sum = 0
    
    For i = 2 To 10 Step 2
        sum = sum + i
    Next i
    
    Range("A1").Value = sum
    
End Sub
