Attribute VB_Name = "Module28"

Sub 参照渡し()

    Dim str As String
    
    str = "伊藤さん"
    
    Call createString(str)
    
    Range("A1").Value = str
    
End Sub

Sub createString(ByRef str As String)
    
    str = str & "、こんにちは"
    
End Sub
