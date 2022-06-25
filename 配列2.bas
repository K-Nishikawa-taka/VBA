Attribute VB_Name = "Module21"

Sub 配列2()

    Dim pref(3 To 6) As String
    Dim i As Integer
    
    Dim src As String
    Dim msg As String
    
    pref(3) = "東京都"
    pref(4) = "神奈川県"
    pref(5) = "千葉県"
    pref(6) = "埼玉県"
    
    src = "埼玉県"
    msg = "関東以外の県です"
    
    For i = 3 To 6
    
        If src = pref(i) Then
            msg = "関東の県です"
        End If
    
    Next i
    
    Range("A1").Value = msg

End Sub
