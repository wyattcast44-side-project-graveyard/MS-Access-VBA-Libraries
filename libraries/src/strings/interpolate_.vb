Attribute VB_Name = "interpolate_"
Option Compare Database
Option Explicit

Public Function interpolateSql(base As String, ParamArray terms()) As String

    Dim term, currentNumber, currentIdentifier, returnString As Variant
    
    currentNumber = 1
    returnString = base
    
    For Each term In terms
    
        currentIdentifier = ":" & CStr(currentNumber) & ":"
        
        Select Case True
            Case VarType(term) = vbInteger
                returnString = Replace(returnString, currentIdentifier, term)
            Case VarType(term) = vbLong
                returnString = Replace(returnString, currentIdentifier, term)
            Case VarType(term) = vbDouble
                returnString = Replace(returnString, currentIdentifier, term)
            Case VarType(term) = vbSingle
                returnString = Replace(returnString, currentIdentifier, term)
            Case VarType(term) = vbDecimal
                returnString = Replace(returnString, currentIdentifier, term)
            Case VarType(term) = vbString
                returnString = Replace(returnString, currentIdentifier, "'" & term & "'")
            Case VarType(term) = vbDate
                returnString = Replace(returnString, currentIdentifier, "#" & term & "#")
            Case VarType(term) = vbBoolean
                returnString = Replace(returnString, currentIdentifier, term)
            Case Else
                returnString = Replace(returnString, currentIdentifier, "'" & term & "'")
        End Select
        
        currentNumber = currentNumber + 1
        
    Next
    
    interpolateSql = returnString

End Function

Public Function interpolate(base As String, ParamArray terms()) As String

    Dim term, currentNumber, currentIdentifier, returnString As Variant
    
    currentNumber = 1
    returnString = base
    
    For Each term In terms
    
        currentIdentifier = ":" & CStr(currentNumber) & ":"
        
        returnString = Replace(returnString, currentIdentifier, term)
        
        currentNumber = currentNumber + 1
        
    Next
    
    interpolate = returnString

End Function

