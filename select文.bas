Attribute VB_Name = "Module12"

Sub select��()

Dim x As String

x = "����"

Select Case x
Case "����"
    Range("A1").Value = "Tokyo"
Case "���"
    Range("A1").Value = "Osaka"
Case "���É�"
    Range("A1").Value = "Nagoya"
Case "����"
    Range("A1").Value = "Fukuoka"
End Select

End Sub
