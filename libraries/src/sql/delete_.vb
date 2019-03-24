Attribute VB_Name = "delete_"
Option Compare Database
Option Explicit

Public Function delete(table As String, identifier As Variant, Optional columnName As String = "") As Boolean

On Error GoTo handleError

    Dim qry As String
    
    If columnName <> "" Then
        
        If IsNumeric(identifier) Then
            qry = "DELETE * FROM " & table & " WHERE(" & columnName & " = " & identifier & ")"
        Else
            qry = "DELETE * FROM " & table & " WHERE(" & columnName & " = '" & CStr(identifier) & "')"
        End If
        
    Else
        
        If IsNumeric(identifier) Then
            qry = "DELETE * FROM " & table & " WHERE(ID = " & identifier & ")"
        Else
            qry = "DELETE * FROM " & table & " WHERE(ID = '" & CStr(identifier) & "')"
        End If
    
    End If
    
    CurrentDb.Execute qry, dbFailOnError
    delete = True
    Exit Function
    
handleError:
    delete = False
    Exit Function
    
End Function
