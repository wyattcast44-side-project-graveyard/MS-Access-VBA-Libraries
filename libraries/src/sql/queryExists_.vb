Attribute VB_Name = "queryExists_"
Option Compare Database
Option Explicit

Public Function queryExists(qryName As String) As Boolean
    
    Dim qdf, cName, toCheck

    queryExists = False

    For Each qdf In CurrentDb.QueryDefs
        cName = qdf.Name
        If cName = qryName Then
            queryExists = True
        End If
    Next
    
End Function

