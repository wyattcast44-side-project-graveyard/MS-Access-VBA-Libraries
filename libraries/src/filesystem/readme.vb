Attribute VB_Name = "readme"
Option Compare Database
Option Explicit

' SOURCE
' - https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'
' METHODS
' - fileExists
' - folderExists
'
' METHOD LIFECYCLES
' - handleError
' - handleSuccess
'
' ERRORS
' - Each method in the library handles errors by calling the handleError() method, you are free to hook into
'   this handleError() method to handle all errors however you want :)

