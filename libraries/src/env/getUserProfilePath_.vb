Attribute VB_Name = "getUserProfilePath_"
Option Compare Database
Option Explicit

Public Function getUserProfilePath(Optional endWithSlash As Boolean = True)
    getUserProfilePath = Environ("USERPROFILE") & IIf(endWithSlash, "\", "")
End Function
