Attribute VB_Name = "Module10"

Sub ������������()
Dim x As String

x = "���"

If x = "����" Then
    Range("A1").Value = "���Z�܂��͓����ł�"
ElseIf x = "���" Then
    Range("A1").Value = "���Z�܂��͑��ł�"
ElseIf x = "���É�" Then
    Range("A1").Value = "���Z�܂��͖��É��ł�"
Else
    Range("A1").Value = "���Z�܂��͂킩��܂���"
End If

End Sub
