Attribute VB_Name = "Module13"

Sub select文2()

Dim x As Integer

x = 10

Select Case x
Case Is < 5
    Range("A1").Value = "xは5より小さい"
Case Is >= 20
    Range("A1").Value = "xは20以上"
Case Else
    Range("A1").Value = "xは5以上20未満"
End Select

End Sub
