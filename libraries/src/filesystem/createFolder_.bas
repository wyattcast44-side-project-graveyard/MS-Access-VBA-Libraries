Option Compare Database
Option Explicit

Public Function createFolder(path As String, Optional failIfAlreadyExists As Boolean = False) As Boolean

On Error GoTo handleError
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path <> "" Then
        FSO.createFolder path
    End If
    
    If folderExists(path) Then
        createFolder = True
    Else
        createFolder = False
    End If
    
    GoTo handleSuccess
    Exit Function

handleSuccess:
    GoTo cleanUp
    Exit Function

handleError:
    If Err.Number = 58 And Not failIfAlreadyExists Then
        createFolder = True
    Else
        Call fileSystem.handleError(Err.Number, Err.Description, "createFolder()", path)
    End If
    GoTo cleanUp
    
cleanUp:
    Set FSO = Nothing
    Exit Function
    
End Function

Public Function test()
    Debug.Print "Custom cb"
End Function 