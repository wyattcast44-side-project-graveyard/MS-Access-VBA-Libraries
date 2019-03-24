Attribute VB_Name = "dexists_"
Option Compare Database
Option Explicit

Public Function dexists(field As String, table As String, criteria As String) As Boolean
    
    dexists = IIf(DCount(field, table, criteria) > 0, True, False)
    
End Function
