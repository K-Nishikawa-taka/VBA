Attribute VB_Name = "Module13"

Sub selectï∂2()

Dim x As Integer

x = 10

Select Case x
Case Is < 5
    Range("A1").Value = "xÇÕ5ÇÊÇËè¨Ç≥Ç¢"
Case Is >= 20
    Range("A1").Value = "xÇÕ20à»è„"
Case Else
    Range("A1").Value = "xÇÕ5à»è„20ñ¢ñû"
End Select

End Sub
