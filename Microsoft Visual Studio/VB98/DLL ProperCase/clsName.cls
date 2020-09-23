VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public variables visible to user when Referenced
Public FirstName As String
Public LastName As String

'Private variables are not visible
Private strPrivateVar As String

'Public subs and function are visible to user
Public Sub EnterName()
Dim str1st As String
Dim str2nd As String

str1st = InputBox("Enter 1st name.", "Dll Tester", , 3000, 2000)
str2nd = InputBox("Enter Last name.", "Dll Tester", , 3000, 2000)

FirstName = ProperCaseName(str1st)
LastName = ProperCaseName(str2nd)

'MsgBox "Your name in Proper Case is " & FirstName & " " & LastName

End Sub

'Private subs and functions are not visible
Private Function ProperCaseName(ByVal strName As String) As String

ProperCaseName = StrConv(strName, vbProperCase)

End Function
