Attribute VB_Name = "Module20"

Sub 配列()

    Dim pref(3) As String
    Dim i As Integer
    
    Dim src As String
    Dim msg As String
    
    pref(0) = "東京都"
    pref(1) = "神奈川県"
    pref(2) = "千葉県"
    pref(3) = "埼玉県"
    
    src = "茨城県"
    msg = "関東以外の県です"
    
    For i = 0 To 3
    
        If src = pref(i) Then
            msg = "関東の県です"
        End If
        
    Next i
    
    Range("A1").Value = msg
    
    src = "埼玉県"
    msg = "関東以外の県です"
    
    For i = 0 To 3
    
        If src = pref(i) Then
            msg = "関東の県です"
        End If
        
    Next i
     
    Range("A2").Value = msg

End Sub
