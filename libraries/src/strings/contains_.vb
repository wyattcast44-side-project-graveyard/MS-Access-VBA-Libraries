Attribute VB_Name = "contains_"
Option Compare Database
Option Explicit

''''''''''''''''''''''''''''
'
' Name:         contains()
' Library:      Strings.accda
' Author:       Wyatt Castaneda
' Last Update:  23-Mar-19
' Description:  Searchs an arbitary number of strings for a substring
'
' Example(s):   contains("wyatt", "wyatt", "james", "amber") --> true
'               contains("scott", "wyatt", "james", "amber") --> false
'
''''''''''''''''''''''''''''

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
