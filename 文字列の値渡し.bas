Attribute VB_Name = "Module26"

Sub 文字列の値渡し()

    Dim str As String
    
    str = "こんにちは"
    
    Call setCellValue(str)
    
    Range("A2").Value = str
    
End Sub

Sub setCellValue(ByVal str As String)
    
    str = str & "お元気ですか"
    Range("A1").Value = str
    
End Sub
