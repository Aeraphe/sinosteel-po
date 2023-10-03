Attribute VB_Name = "db_users"


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
    On Error Resume Next
    Create = database.Insert("users", data)

End Function


Public Function get_role(user_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT RO.name As role FROM  users AS US INNER JOIN user_roles AS RO ON RO.id = US.role_id  Where  US.id  = " & user_id


    Set get_role = database.cn.Execute(sqlStrQuery)


End Function


Public Function getUsersBySector(sector As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT US.id,US.name FROM  users AS US INNER JOIN user_roles AS RO ON RO.id = US.role_id  Where  RO.sector  ='" & sector & "'"


    Set getUsersBySector = database.cn.Execute(sqlStrQuery)


End Function
