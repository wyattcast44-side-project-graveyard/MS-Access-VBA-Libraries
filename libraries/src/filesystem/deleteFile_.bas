Option Compare Database
Option Explicit

Public Function deleteFile(path As String) As Boolean

On Error GoTo handleError
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path <> "" And fileExists(path) Then
        FSO.deleteFile path
    Else
        Exit Function
    End If
    
    If fileExists(path) Then
        deleteFile = False
    Else
        deleteFile = True
    End If

    GoTo handleSuccess
    Exit Function

handleSuccess:
    Call fileSystem.handleSuccess
    GoTo cleanUp
    Exit Function
    
handleError:
    Call fileSystem.handleError(Err.Number, Err.Description, "deleteFolder()", path)
    GoTo cleanUp
    Exit Function
    
cleanUp:
    Set FSO = Nothing
    Exit Function

End Function 