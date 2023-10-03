Attribute VB_Name = "database_equipaments"


'namespace=vba-files\DataBase




Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("projects_equipaments")

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

    sqlStrQuery = " DELETE   FROM  projects_equipaments  Where id = " & idData

    database.cn.Execute (sqlStrQuery)



End Function



Public Function getEquipamentsFromProject(ByVal project_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT  * FROM  projects_equipaments  Where  project_id  = " & project_id & "  ORDER BY id DESC "


    Set getEquipamentsFromProject = database.cn.Execute(sqlStrQuery)


End Function
