Attribute VB_Name = "db_doc_store"


'namespace=vba-files\DataBase




Public Function getProjectDocumentFolders(doc_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT  * FROM  project_document_folders  Where  project_id  = " & doc_id & "  ORDER BY id DESC "


    Set getProjectDocumentFolders = database.cn.Execute(sqlStrQuery)


End Function
