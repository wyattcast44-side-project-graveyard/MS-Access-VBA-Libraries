Attribute VB_Name = "ucFirst_"
Option Compare Database
Option Explicit

Public Function ucFirst(toFix As String) As String
    On Error GoTo failGracefully
        ucFirst = StrConv(Left(toFix, 1), vbUpperCase) & Right(toFix, Len(toFix) - 1)
        Exit Function
failGracefully:
    ucFirst = toFix
    Exit Function
End Function
