Attribute VB_Name = "Module10"

Sub 複数条件分岐()
Dim x As String

x = "大阪"

If x = "東京" Then
    Range("A1").Value = "お住まいは東京です"
ElseIf x = "大阪" Then
    Range("A1").Value = "お住まいは大阪です"
ElseIf x = "名古屋" Then
    Range("A1").Value = "お住まいは名古屋です"
Else
    Range("A1").Value = "お住まいはわかりません"
End If

End Sub
