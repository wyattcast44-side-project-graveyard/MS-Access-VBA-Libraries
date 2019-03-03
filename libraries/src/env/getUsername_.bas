Option Compare Database
Option Explicit

Public Function getCurrentUsername()
    getCurrentUsername = Environ("USERNAME")
End Function
 