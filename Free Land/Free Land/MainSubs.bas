Attribute VB_Name = "MainSubs"

Global LanguageTx(200) As String
Dim LanguageTxNum As Integer

Global I As Integer
Global I2 As Integer
Global x As String
Global x2 As String
Global x3 As String
Global x4 As String
Global x5 As String
Global x6 As String
Global x7 As String
Global x8 As String
Global x9 As String


Function LoadLanguage(Arb As Boolean)
If Arb = True Then
Open App.Path + "\Data\Languages\Arabic.Txt" For Input As #1
Else
Open App.Path + "\Data\Languages\English.Txt" For Input As #1
End If

On Error GoTo 1
LanguageTxNum = 0
For I = 0 To 200
Input #1, x
LanguageTx(LanguageTxNum) = x
LanguageTxNum = LanguageTxNum + 1
Next
1:
Close #1
End Function
