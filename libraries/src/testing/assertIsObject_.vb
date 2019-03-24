Attribute VB_Name = "assertIsObject_"
Option Compare Database
Option Explicit

Public Function assertIsObject(item As Variant)
    assertIsObject = IIf(IsObject(item), True, False)
End Function
