Attribute VB_Name = "Module13"

Sub select��2()

Dim x As Integer

x = 10

Select Case x
Case Is < 5
    Range("A1").Value = "x��5��菬����"
Case Is >= 20
    Range("A1").Value = "x��20�ȏ�"
Case Else
    Range("A1").Value = "x��5�ȏ�20����"
End Select

End Sub
