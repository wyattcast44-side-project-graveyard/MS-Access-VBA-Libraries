Attribute VB_Name = "lcFirst_"
Option Compare Database
Option Explicit

Public Function lcFirst(toFix As String) As String
    On Error GoTo failGracefully
        lcFirst = StrConv(Left(toFix, 1), vbLowerCase) & Right(toFix, Len(toFix) - 1)
        Exit Function
failGracefully:
    lcFirst = toFix
    Exit Function
End Function
