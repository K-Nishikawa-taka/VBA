Attribute VB_Name = "Module11"

Sub 多重条件分岐()

Dim x As Integer

x = 9

If x > 0 Then
    If x Mod 2 = 0 Then
        Range("A1").Value = "xは正の偶数です"
    Else
        Range("A1").Value = "xは正の奇数です"
    End If
Else
    Range("A1").Value = "xは負の数です"
End If

End Sub
