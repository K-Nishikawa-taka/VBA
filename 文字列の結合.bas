Attribute VB_Name = "Module6"

Sub 文字列の結合()
Dim str1 As String
Dim str2 As String
Dim num1 As Integer
Dim num2 As Integer

str1 = "こんにちは"
str2 = "お元気ですか"
num1 = 10
num2 = 34

Range("A1").Value = str1 & str2
Range("A2").Value = str1 & num1
Range("A3").Value = num1 & str2
Range("A4").Value = num1 & num2
End Sub
