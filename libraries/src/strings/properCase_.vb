Attribute VB_Name = "properCase_"
Option Compare Database
Option Explicit

Public Function properCase(toFix As String) As String
    On Error GoTo failGracefully
        properCase = StrConv(toFix, vbProperCase)
        Exit Function
failGracefully:
    properCase = toFix
    Exit Function
End Function
