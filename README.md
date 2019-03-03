# MS Access Libraries

This is a collection on standalone libraries for MS Access projects. This is a WIP, I do not suggesting using in any sort of production project at this time.

## Installation

1. Download the library(ies) that you want to use
2. Create a `libraries` folder in your project root 
3. Add the `addReferences` function below to your main project
4. Ensure that this `addReferences` function is called when you open the database

```vba
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
```

## Usage

To use any public function defined in the libaries, simpy call them. If there may be a case where you have the same function defined in your main project, call the function specifically from the library, for example `libraryName.functionName` 

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
[MIT](https://choosealicense.com/licenses/mit/)