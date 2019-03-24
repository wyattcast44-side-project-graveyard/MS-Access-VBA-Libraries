Attribute VB_Name = "assertStringEquals_"
Option Compare Database
Option Explicit

Public Function assertStringEquals(value As String, correctValue As String, Optional caseSensitive As Boolean = True)
    assertStringEquals = IIf(StrComp(value, correctValue, IIf(caseSensitive, vbBinaryCompare, vbTextCompare)) = 0, True, False)
End Function

