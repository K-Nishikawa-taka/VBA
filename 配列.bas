Attribute VB_Name = "Module20"

Sub �z��()

    Dim pref(3) As String
    Dim i As Integer
    
    Dim src As String
    Dim msg As String
    
    pref(0) = "�����s"
    pref(1) = "�_�ސ쌧"
    pref(2) = "��t��"
    pref(3) = "��ʌ�"
    
    src = "��錧"
    msg = "�֓��ȊO�̌��ł�"
    
    For i = 0 To 3
    
        If src = pref(i) Then
            msg = "�֓��̌��ł�"
        End If
        
    Next i
    
    Range("A1").Value = msg
    
    src = "��ʌ�"
    msg = "�֓��ȊO�̌��ł�"
    
    For i = 0 To 3
    
        If src = pref(i) Then
            msg = "�֓��̌��ł�"
        End If
        
    Next i
     
    Range("A2").Value = msg

End Sub
