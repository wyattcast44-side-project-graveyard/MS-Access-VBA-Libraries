Option Compare Database
Option Explicit

Public Function lowerCase(toFix As String) As String
    On Error GoTo failGracefully
        lowerCase = StrConv(toFix, vbLowerCase)
        Exit Function
failGracefully:
    lowerCase = toFix
    Exit Function
End Function 