Attribute VB_Name = "Module11"

Sub ���d��������()

Dim x As Integer

x = 9

If x > 0 Then
    If x Mod 2 = 0 Then
        Range("A1").Value = "x�͐��̋����ł�"
    Else
        Range("A1").Value = "x�͐��̊�ł�"
    End If
Else
    Range("A1").Value = "x�͕��̐��ł�"
End If

End Sub
