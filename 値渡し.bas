Attribute VB_Name = "Module25"

Sub 値渡し()
    Call warikireCheck(10)
End Sub

Sub warikireCheck(ByVal num As Integer)
    If num Mod 2 = 0 Then
        Range("A1").Value = "割り切れます"
    Else
        Range("A1").Value = "割り切れません"
    End If
End Sub
