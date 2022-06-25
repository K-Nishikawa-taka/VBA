Attribute VB_Name = "Module14"
Sub ƒ‹[ƒv()

Dim sum As Integer
Dim x As Integer

sum = 0
x = 1

Do While x <= 10
    sum = sum + x
    x = x + 1
Loop

Range("A1").Value = sum

End Sub
