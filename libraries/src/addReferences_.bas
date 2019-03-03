Attribute VB_Name = "addReferences_"
Option Compare Database
Option Explicit

Public Function addReferences()

On Error Resume Next

    Dim fileObj, FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    For Each fileObj In FSO.GetFolder(Application.CurrentProject.path & "\libraries").files
        If FSO.GetExtensionName(fileObj.path) = "accda" Then
            Application.References.AddFromFile fileObj.path
        End If
    Next
    
End Function

