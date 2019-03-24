Attribute VB_Name = "getHTTPClient_"
Option Compare Database
Option Explicit

Public Function getHTTPClient()

On Error GoTo handleError

    Dim client As Object
    
    Set client = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Set getHTTPClient = client
    Exit Function

handleError:
    Set getHTTPClient = Nothing
    Exit Function
    
End Function
