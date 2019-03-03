Attribute VB_Name = "getHomePath_"
Option Compare Database
Option Explicit

Public Function getHomePath(Optional includeDrive As Boolean = True, Optional endWithSlash As Boolean = True)
    getHomePath = IIf(includeDrive, getHomeDrive, "") & Environ("HOMEPATH") & IIf(endWithSlash, "\", "")
End Function
