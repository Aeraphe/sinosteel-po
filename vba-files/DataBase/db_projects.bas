Attribute VB_Name = "db_projects"


'namespace=vba-files\DataBase


Public Function get_by_id(ByVal project_id As String) As ADODB.Recordset


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT * FROM  projects  Where id ='" & project_id & "' ORDER BY id DESC "
    Set get_by_id = database.cn.Execute(sqlStrQuery)


End Function



Public Function get_contract_items(ByVal project_id As String) As ADODB.Recordset


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT * FROM  project_contract_items  Where project_id ='" & project_id & "' ORDER BY id DESC "
    Set get_contract_items = database.cn.Execute(sqlStrQuery)


End Function
