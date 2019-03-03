Attribute VB_Name = "exportModules_"
Option Compare Database

Public Function exportModules()

On Error GoTo handleError
    
    Dim fileSystem, moduleObj, fileObj As Object
    Dim path, moduleCount, moduleName
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = CurrentProject.path & "\src\" & LCase(Replace(CurrentProject.Name, ".accda", "")) & "\"
    
    '' Delete all old files
    Kill path & "*.bas"
    
    moduleCount = Access.CurrentProject.AllModules.Count

    For i = 0 To (moduleCount) - 1
        moduleName = Application.CurrentProject.AllModules(i).Name
        Access.DoCmd.OpenModule moduleName
        Set moduleObj = Access.Modules(moduleName)
        If moduleObj.CountOfLines > 0 Then
           Set fileObj = fileSystem.CreateTextFile(path & moduleName & ".bas")
           fileObj.write moduleObj.Lines(1, moduleObj.CountOfLines) & " "
        End If
        Access.DoCmd.Close acModule, moduleName
        Set moduleObj = Nothing
    Next
    
    Set fileObj = Nothing
    Set fileSystem = Nothing
    
    Exit Function

handleError:
    Debug.Print Err.Number
    Debug.Print Err.Description
    Exit Function

End Function

