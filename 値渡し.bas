Attribute VB_Name = "Module25"

Sub ’l“n‚µ()
    Call warikireCheck(10)
End Sub

Sub warikireCheck(ByVal num As Integer)
    If num Mod 2 = 0 Then
        Range("A1").Value = "Š„‚èØ‚ê‚Ü‚·"
    Else
        Range("A1").Value = "Š„‚èØ‚ê‚Ü‚¹‚ñ"
    End If
End Sub
