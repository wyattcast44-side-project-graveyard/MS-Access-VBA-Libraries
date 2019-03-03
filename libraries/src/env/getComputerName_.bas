Option Compare Database
Option Explicit

Public Function getComputerName()
    getComputerName = Environ("COMPUTERNAME")
End Function 