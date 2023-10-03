Attribute VB_Name = "db_grd_recipient"


'namespace=vba-files\DataBase


'/*
'
'Get last Recipent
'
'*/
Public Function get_by_id(ByVal id As String, ByVal project_id As String) As Variant

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT  REC.id,REC.user_id,REC.email_msg_id, REC.supplier_id,REC.name,REC.code,REC.folder_name  FROM  grd_recipients AS REC   Where project_id = " & project_id & " AND id=" & id

    Set get_by_id = database.cn.Execute(sqlStrQuery)



End Function

'/*
'
'Get last Recipent
'
'*/
Public Function get_all_recipent_from_project(ByRef project_id As String) As Variant

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT REC.id, REC.supplier_id,REC.name,REC.code,REC.folder_name  FROM  grd_recipients AS REC   Where project_id = " & project_id & " ORDER BY id DESC "

    Set get_all_recipent_from_project = database.cn.Execute(sqlStrQuery)



End Function

'/*
'
'Get all emails from recipient
'
'*/
Public Function get_all_recipent_emails(ByRef recipient_id As String) As Variant

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT *  FROM  grd_recipient_emails    Where recipient_id = " & recipient_id & " ORDER BY id DESC "

    Set get_all_recipent_emails = database.cn.Execute(sqlStrQuery)



End Function


'/*
'
'Delete Recipient By ID
'
'*/
Public Function delete_recipient(ByVal id As String, ByVal project_id As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  grd_recipients  Where id = " & id & " AND project_id = " & project_id

    database.cn.Execute (sqlStrQuery)


End Function


'/*
'
'Delete Recipient email
'
'*/
Public Function delete_recipient_email(ByVal id As String, ByVal recipient_id As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  grd_recipient_emails  Where recipient_id = '" & recipient_id & "' AND id='" & id & "'"

    database.cn.Execute (sqlStrQuery)


End Function



Public Function add_email(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    add_email = database.Insert("grd_recipient_emails", data)

End Function


Public Function Create(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    CreateGRD = database.Insert("grd_recipient", data)

End Function



Public Function update(ByVal data As Variant, ByVal where As String)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Call database.update("grd_recipients", data, where)

End Function

