Attribute VB_Name = "Module17"
Sub Loop•¶()
    Dim x As Integer
    
    x = 1
    
    Do
        x = x * 3
    Loop While x < 100
    
    Range("A1").Value = x
    
End Sub
