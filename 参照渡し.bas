Attribute VB_Name = "Module28"

Sub �Q�Ɠn��()

    Dim str As String
    
    str = "�ɓ�����"
    
    Call createString(str)
    
    Range("A1").Value = str
    
End Sub

Sub createString(ByRef str As String)
    
    str = str & "�A����ɂ���"
    
End Sub
