Attribute VB_Name = "contains_"
Option Compare Database
Option Explicit

Public Function contains(toCheck As String, ParamArray searchTerms()) As Boolean
        
    Dim term As Variant
    
    contains = False
    
    For Each term In searchTerms
        If InStr(toCheck, term) <> 0 Then
            GoTo doesContainString
        End If
    Next
    
    Exit Function
       
doesContainString:
    contains = True
    Exit Function
    
End Function
