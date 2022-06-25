Attribute VB_Name = "Module12"

Sub select•¶()

Dim x As String

x = "•Ÿ‰ª"

Select Case x
Case "“Œ‹"
    Range("A1").Value = "Tokyo"
Case "‘åã"
    Range("A1").Value = "Osaka"
Case "–¼ŒÃ‰®"
    Range("A1").Value = "Nagoya"
Case "•Ÿ‰ª"
    Range("A1").Value = "Fukuoka"
End Select

End Sub
