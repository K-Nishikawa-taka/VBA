Attribute VB_Name = "Module12"

Sub selectถ()

Dim x As String

x = "ช"

Select Case x
Case ""
    Range("A1").Value = "Tokyo"
Case "ๅใ"
    Range("A1").Value = "Osaka"
Case "ผรฎ"
    Range("A1").Value = "Nagoya"
Case "ช"
    Range("A1").Value = "Fukuoka"
End Select

End Sub
