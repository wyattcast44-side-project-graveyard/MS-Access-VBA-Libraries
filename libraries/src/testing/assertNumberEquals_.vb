Attribute VB_Name = "assertNumberEquals_"
Option Compare Database
Option Explicit

Public Function assertNumberEquals(value As Variant, correctValue As Variant)
    If IsNumeric(value) And IsNumeric(correctValue) Then
        assertNumberEquals = IIf(value = correctValue, True, False)
    Else
        assertNumberEquals = False
    End If
End Function
