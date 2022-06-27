Attribute VB_Name = "Module26"

Sub •¶Žš—ñ‚Ì’l“n‚µ()

    Dim str As String
    
    str = "‚±‚ñ‚É‚¿‚Í"
    
    Call setCellValue(str)
    
    Range("A2").Value = str
    
End Sub

Sub setCellValue(ByVal str As String)
    
    str = str & "‚¨Œ³‹C‚Å‚·‚©"
    Range("A1").Value = str
    
End Sub
