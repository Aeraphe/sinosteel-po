Attribute VB_Name = "ProjectDataBase"


'namespace=vba-files\DataBase

'/*
'
'Create Sppliers
'
'@param <Array>  data
'
'*/
Public Function Create(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    'Insert New Budget
    Dim id As Long

    id = database.Insert("projects", data)

End Function


Public Function getAll(Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "name") As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("projects", searchWord, filterType)

End Function
