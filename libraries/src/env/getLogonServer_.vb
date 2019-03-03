Attribute VB_Name = "getLogonServer_"
Option Compare Database
Option Explicit

Public Function getLogonServer()
    getLogonServer = Environ("LOGONSERVER")
End Function
