Attribute VB_Name = "getRSCount_"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : getRSCount
' Author    : Wyatt Castaneda
' Date      : 11/02/2018
' Purpose   : Takes a given recordset and return the count of records in said recordset
' Params    : rs as recordset
' Returns   : interger
' Test      : none
'---------------------------------------------------------------------------------------

Public Function getRSCount(rs As Recordset) As Integer

On Error GoTo handleError
    Dim count As Integer
    
    count = 0
    
    If Not rs.EOF Then
        
        rs.MoveFirst
        rs.MoveLast
        
        count = rs.RecordCount
        
    End If
    
    getRSCount = count
    Exit Function
    
handleError:
    getRSCount = count
    Exit Function

End Function
