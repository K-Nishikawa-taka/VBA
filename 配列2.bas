Attribute VB_Name = "Module21"

Sub �z��2()

    Dim pref(3 To 6) As String
    Dim i As Integer
    
    Dim src As String
    Dim msg As String
    
    pref(3) = "�����s"
    pref(4) = "�_�ސ쌧"
    pref(5) = "��t��"
    pref(6) = "��ʌ�"
    
    src = "��ʌ�"
    msg = "�֓��ȊO�̌��ł�"
    
    For i = 3 To 6
    
        If src = pref(i) Then
            msg = "�֓��̌��ł�"
        End If
    
    Next i
    
    Range("A1").Value = msg

End Sub
