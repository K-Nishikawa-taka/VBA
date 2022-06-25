Attribute VB_Name = "Module11"

Sub ‘½dğŒ•ªŠò()

Dim x As Integer

x = 9

If x > 0 Then
    If x Mod 2 = 0 Then
        Range("A1").Value = "x‚Í³‚Ì‹ô”‚Å‚·"
    Else
        Range("A1").Value = "x‚Í³‚ÌŠï”‚Å‚·"
    End If
Else
    Range("A1").Value = "x‚Í•‰‚Ì”‚Å‚·"
End If

End Sub
