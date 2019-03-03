Attribute VB_Name = "install_"
Option Compare Database
Option Explicit

Const tableName As String = "tblPermissionsManagement"

Public Function install(database)
    
    Call createTables(database)
    
End Function

Public Function createTables(database)
    
    Dim qry As String
    qry = "CREATE TABLE " & tableName & " (ID COUNTER(1, 1) PRIMARY KEY, itemName VARCHAR(255), itemDesc VARCHAR(255), accessLevelRequired SMALLINT, created_at DATETIME, updated_at DATETIME)"
    database.Execute qry, dbFailOnError
    
End Function

Public Function authorize(user, action)
    ''
End Function

