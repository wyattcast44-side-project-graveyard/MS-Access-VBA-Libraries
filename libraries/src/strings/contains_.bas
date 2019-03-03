Option Compare Database
Option Explicit

Public Function contains(toCheck As String, searchTerm As String) As Boolean
    On Error GoTo failGracefully
        contains = IIf(InStr(toCheck, searchTerm) <> 0, True, False)
        Exit Function
failGracefully:
    contains = False
    Exit Function
End Function 