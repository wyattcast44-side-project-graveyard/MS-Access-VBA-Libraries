Option Compare Database
Option Explicit

Public Function truncate(originalStr As String, length As Integer, Optional paddingString = "") As String
    
    truncate = Trim(Left(originalStr, length) & paddingString)
    
End Function 