Attribute VB_Name = "Module25"

Sub �l�n��()
    Call warikireCheck(10)
End Sub

Sub warikireCheck(ByVal num As Integer)
    If num Mod 2 = 0 Then
        Range("A1").Value = "����؂�܂�"
    Else
        Range("A1").Value = "����؂�܂���"
    End If
End Sub
