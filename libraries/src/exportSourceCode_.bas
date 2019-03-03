Attribute VB_Name = "exportSourceCode_"
Option Compare Database

Public Function exportSourceCode()
    
    Dim basePath, projectPath, exportPath, filePath As Variant
    Dim library, libraries As Variant
    Dim accessObj As Access.Application
    Dim Module As VBComponent
    
    DoCmd.Hourglass True
    
    basePath = CurrentProject.path & "\"
    
    exportPath = basePath & "src\"
    
    libraries = Array( _
        "Env.accda", _
        "Strings.accda" _
    )
    
    Set accessObj = New Access.Application
    
    For Each library In libraries
        
        exportPath = basePath & "src\"
    
        projectPath = basePath & library
        
        accessObj.OpenCurrentDatabase (projectPath)
        
        exportPath = exportPath & LCase(Replace(accessObj.Application.CurrentProject.Name, ".accda", "")) & "\"
        
        Kill exportPath & "*.*"
        
        For Each Module In accessObj.VBE.ActiveVBProject.VBComponents
            
            Module.Export exportPath & Module.Name & ".vb"
            
        Next
        
        accessObj.CloseCurrentDatabase
    
    Next
    
    Set accessObj = Nothing
    
    DoCmd.Hourglass False
    
End Function

