Attribute VB_Name = "db_documents_discipline"


'namespace=vba-files\DataBase


Public Function getAll(id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT * FROM  discipline  Where  id  = " & id & "  ORDER BY id DESC "


    Set getAll = database.cn.Execute(sqlStrQuery)


End Function


    
