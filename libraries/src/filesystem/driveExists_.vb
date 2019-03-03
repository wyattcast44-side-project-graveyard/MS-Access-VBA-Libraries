Attribute VB_Name = "driveExists_"
Option Compare Database
Option Explicit

Public Function driveExists(path As String) As Boolean

On Error GoTo handleError
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path <> "" Then
        driveExists = FSO.driveExists(path)
    End If

    GoTo handleSuccess
    Exit Function

handleSuccess:
    Call fileSystem.handleSuccess
    GoTo cleanUp
    Exit Function

handleError:
    Call fileSystem.handleError(Err.Number, Err.Description, "driveExists()", path)
    GoTo cleanUp

cleanUp:
    Set FSO = Nothing
    Exit Function

End Function
