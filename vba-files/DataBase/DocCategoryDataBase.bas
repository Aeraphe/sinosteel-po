Attribute VB_Name = "DocCategoryDataBase"


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

    Create = database.Insert("document_categories", data)

End Function


Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("document_categories")

End Function



'
'Delete budget By ID
'
'*/
Public Function delete(ByRef idData As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  document_categories  Where id = " & idData

    database.cn.Execute (sqlStrQuery)



End Function
