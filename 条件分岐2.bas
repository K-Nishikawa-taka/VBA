Attribute VB_Name = "Module9"

Sub 条件分岐2()
Dim x As Integer

x = 12
y = 8

If x > 10 Then
    Range("A1").Value = "xは10より大きい"
Else
    Range("A1").Value = "xは10より小さい"
End If

If y > 10 Then
    Range("A2").Value = "yは10より大きい"
Else
    Range("A2").Value = "yは10より小さい"
End If

End Sub

