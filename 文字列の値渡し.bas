Attribute VB_Name = "Module26"

Sub ������̒l�n��()

    Dim str As String
    
    str = "����ɂ���"
    
    Call setCellValue(str)
    
    Range("A2").Value = str
    
End Sub

Sub setCellValue(ByVal str As String)
    
    str = str & "�����C�ł���"
    Range("A1").Value = str
    
End Sub
