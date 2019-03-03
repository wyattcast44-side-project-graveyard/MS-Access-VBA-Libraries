Attribute VB_Name = "deleteFolder_"
Option Compare Database
Option Explicit

Public Function deleteFolder(path As String) As Boolean

On Error GoTo handleError
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path <> "" And folderExists(path) Then
        path = IIf(Right(path, 1) = "\", Left(path, Len((path)) - 1), path)
        FSO.deleteFolder path
    Else
        Exit Function
    End If
    
    If folderExists(path) Then
        deleteFolder = False
    Else
        deleteFolder = True
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

cleanUp:
    Set FSO = Nothing
    Exit Function

End Function
