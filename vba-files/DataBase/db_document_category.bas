Attribute VB_Name = "db_document_category"


'namespace=vba-files\DataBase


Public Function getAll(id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT * FROM  document_categories  Where  id  = " & id & "  ORDER BY id DESC "


    Set getAll = database.cn.Execute(sqlStrQuery)


End Function



    
