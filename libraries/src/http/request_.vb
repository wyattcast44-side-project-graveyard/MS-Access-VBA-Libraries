Attribute VB_Name = "request_"
Option Compare Database
Option Explicit

Public Function request(url As String, requestVerb As String, Optional async As Boolean = False) As Variant

    Dim client As Object
    Dim response As Variant
    
    Set client = getHTTPClient()
    
    client.Open UCase(Trim(requestVerb)), url, async
    
    Set request = client
    
    ' client.Send
    
    ' response = client.responseText
    
    'Set client = Nothing

End Function

Public Function test()

    Dim client
    
    Set client = request("https://jsonplaceholder.typicode.com/posts", "GET")

End Function
