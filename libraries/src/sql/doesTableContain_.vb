Attribute VB_Name = "doesTableContain_"
Option Compare Database
Option Explicit

Public Function doesTableContain(table As String, searchCriteria As String, searchColumn As String, Optional fuzzyMatch As Boolean = True) As Boolean
    
    Dim selectQry As String
    Dim rs As Recordset
    
    If fuzzyMatch Then
        selectQry = "SELECT * FROM " & table & " WHERE " & searchColumn & " LIKE '*" & searchCriteria & "*'"
    Else
        selectQry = "SELECT * FROM " & table & " WHERE " & searchColumn & " = '" & searchCriteria & "'"
    End If
    
    Set rs = CurrentDb.OpenRecordset(selectQry)
    
    If getRSCount(rs) <> 0 Then
        doesTableContain = True
    Else
        doesTableContain = False
    End If
    
    Set rs = Nothing
    
End Function
