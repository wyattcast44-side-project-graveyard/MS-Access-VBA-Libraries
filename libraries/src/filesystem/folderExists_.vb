Attribute VB_Name = "folderExists_"
Option Compare Database
Option Explicit

Public Function folderExists(path As String) As Boolean

On Error GoTo handleError
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path <> "" Then
        folderExists = FSO.folderExists(path)
    End If

    GoTo handleSuccess
    Exit Function

handleSuccess:
    Call fileSystem.handleSuccess
    GoTo cleanUp
    Exit Function
    
handleError:
    Call fileSystem.handleError(Err.Number, Err.Description, "folderExists()", path)
    GoTo cleanUp

cleanUp:
    Set FSO = Nothing
    Exit Function

End Function
