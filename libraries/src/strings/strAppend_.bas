Option Compare Database
Option Explicit

Public Function strAppend(stringOne As String, stringTwo As String, Optional separator As String = "")
    strAppend = stringOne & separator & stringTwo
End Function 