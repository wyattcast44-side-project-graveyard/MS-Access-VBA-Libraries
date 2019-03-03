Attribute VB_Name = "fileExists_"
Option Compare Database
Option Explicit

Public Function fileExists(path As String) As Boolean

On Error GoTo handleError
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path <> "" Then
        fileExists = FSO.fileExists(path)
    End If

    GoTo handleSuccess
    Exit Function

handleSuccess:
    Call fileSystem.handleSuccess
    GoTo cleanUp
    Exit Function
    
handleError:
    Call fileSystem.handleError(Err.Number, Err.Description, "fileExists()", path)
    GoTo cleanUp
    
cleanUp:
    Set FSO = Nothing
    Exit Function

End Function
