Option Compare Database
Option Explicit

Public Function getHomeDrive()
    getHomeDrive = Environ("HOMEDRIVE")
End Function 