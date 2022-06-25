Attribute VB_Name = "Module10"

Sub •¡”ğŒ•ªŠò()
Dim x As String

x = "‘åã"

If x = "“Œ‹" Then
    Range("A1").Value = "‚¨Z‚Ü‚¢‚Í“Œ‹‚Å‚·"
ElseIf x = "‘åã" Then
    Range("A1").Value = "‚¨Z‚Ü‚¢‚Í‘åã‚Å‚·"
ElseIf x = "–¼ŒÃ‰®" Then
    Range("A1").Value = "‚¨Z‚Ü‚¢‚Í–¼ŒÃ‰®‚Å‚·"
Else
    Range("A1").Value = "‚¨Z‚Ü‚¢‚Í‚í‚©‚è‚Ü‚¹‚ñ"
End If

End Sub
