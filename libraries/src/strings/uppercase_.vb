Attribute VB_Name = "uppercase_"
Option Compare Database
Option Explicit

Public Function upperCase(toFix As String) As String
    On Error GoTo failGracefully
        upperCase = StrConv(toFix, vbUpperCase)
        Exit Function
failGracefully:
    upperCase = toFix
    Exit Function
End Function
