Attribute VB_Name = "getUserDesktopPath_"
Option Compare Database
Option Explicit

Public Function getUserDesktopPath(Optional endWithSlash As Boolean = True)
    getUserDesktopPath = getUserProfilePath & "Desktop" & IIf(endWithSlash, "\", "")
End Function
