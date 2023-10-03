Attribute VB_Name = "db_porject_files"


'namespace=vba-files\DataBase




Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("project_files")

End Function


Public Function get_by_type(ByRef project_id As String, file_type As String) As Variant

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT *  FROM  project_files  Where project_id = " & project_id & " AND file_type='" & file_type & "' ORDER BY id DESC LIMIT 1"

    Set get_by_type = database.cn.Execute(sqlStrQuery)



End Function


