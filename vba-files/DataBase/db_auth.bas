Attribute VB_Name = "db_auth"


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
    Create = database.Insert("doc_flow", data)

End Function



Public Function get_all_roles() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set get_all_roles = database.getAll("user_roles")

End Function


Public Function get_user(login As String, pass As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT US.name,US.login, US.email,US.id,US.phone,RO.name As role FROM  users AS US INNER JOIN user_roles AS RO ON RO.id = US.role_id  Where  login  = '" & login & "'  And  password='" & pass & "'"


    Set get_user = database.cn.Execute(sqlStrQuery)


End Function
