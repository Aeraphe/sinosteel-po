Attribute VB_Name = "db_app_folders"


'namespace=vba-files\DataBase\App




Public Function get_all() As ADODB.Recordset


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT * FROM  app_folders  ORDER BY id DESC "
    Set get_all = database.cn.Execute(sqlStrQuery)


End Function
