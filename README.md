# MS Access Libraries

This is a collection of standalone VBA libraries for MS Access projects. This is a WIP, I do not suggesting using in any sort of production project at this time.

The basic idea is to pull libraries references into your main project programmatically and automatically. To do this, we can save all libraries as `.accda` files, and place them all in a folder called `libraries`, then when the database is opened, we loop through these libraries and add them as references to your project. 

This keeps your main application clean, with only the business logic unique to your project while still benefiting from useful, reusable, libraries of code. This is, in my opinion, a much better option than copying and pasting random code into your project.

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

## Current Libraries

### Auth
The Auth library will provide an flexible, easy to use, authentication system for controlling access to your database.

### Env
The Env library provides easy access to various windows enviroment paths/values.

### Filesystem
The Filesystem library provides easy to use methods for common filesystem tasks such as creating/deleting files/folders.

### Http
The Http library will provide easy to use methods for performing HTTP requests, and handling responses.

### Logger
The Logger library will provide easy to use and configure methods to log errors and activity in your database.

### Middleware
The Middleware library will provide an flexible, easy to use, authorization system for managing roles and permission for your database.

### Outlook
The Outlook library will provide easy to use and configure methods to perform common tasks in MS Outlook.

### SQL
The SQL library will provide easy to use and configure methods to perform SQL tasks (inserts, deletes, etc) in VBA for your database.

### SourceControl
The SourceControl library provides an easy way to export all VBA modules in your database(s) for checking into source control. For an example, the `src` files in this repo were exported from their respective databases using the SourceControl library. This library is mainly to aide in the maintenance of this repo, but you may find a use for it.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Feedback
I'd love to hear your thoughts on this proposed architecture, reach out to me on [twitter](https://twitter.com/WyattCastaned44)

## License
[MIT](https://choosealicense.com/licenses/mit/)

## Good Resources
- [MS Access SQL Reference](https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/microsoft-access-sql-reference)