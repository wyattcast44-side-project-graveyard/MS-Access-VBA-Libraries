Attribute VB_Name = "insert_"
Option Compare Database
Option Explicit

Public Function insert(table As String, columns As String, ParamArray values()) As Variant

On Error GoTo handleError
    
    Dim insertQry, value, valueStr, columnCount, argumentCount As Variant

    valueStr = ""
    argumentCount = (UBound(values) - LBound(values) + 1)
    columnCount = (UBound(Split(columns, ",")) - LBound(Split(columns, ",")) + 1)
    
    If argumentCount <> columnCount Then GoTo handleError
    
    For Each value In values
        
        Select Case True
            Case VarType(value) = vbInteger
                valueStr = valueStr & value & ","
            Case VarType(value) = vbLong
                valueStr = valueStr & value & ","
            Case VarType(value) = vbDouble
                valueStr = valueStr & value & ","
            Case VarType(value) = vbSingle
                valueStr = valueStr & value & ","
            Case VarType(value) = vbDecimal
                valueStr = valueStr & value & ","
            Case VarType(value) = vbString
                valueStr = valueStr & "'" & value & "',"
            Case VarType(value) = vbDate
                valueStr = valueStr & "#" & value & "#,"
            Case VarType(value) = vbBoolean
                valueStr = valueStr & value & ","
            Case Else
                valueStr = valueStr & "'" & CStr(value) & "',"
        End Select
        
    Next
    
    valueStr = IIf(Right(valueStr, 1) = ",", Left(valueStr, (Len(valueStr) - 1)), valueStr)
    
    insertQry = "INSERT INTO " & table & " (" & CStr(columns) & ") VALUES(" & valueStr & ")"
    
    CurrentDb.Execute insertQry, dbFailOnError
    
    insert = True
    
    Exit Function
    
handleError:
    insert = False
    Exit Function

End Function
