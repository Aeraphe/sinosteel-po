Attribute VB_Name = "db_email"


'namespace=vba-files\DataBase




'/*
'
'
'
'*/
Public Function get_config_msgs(ByVal poriject_id As String, ByVal section_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT * FROM  mail_config_msgs  Where project_id ='" & poriject_id & "' AND section_id='" & section_id & "'  ORDER BY id DESC "

    Set get_config_msgs = database.cn.Execute(sqlStrQuery)



End Function


'/*
'
'
'
'*/
Public Function get_email_conf_from_recipient(ByVal id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT MG.id,MG.name FROM grd_recipients AS REC INNER JOIN mail_config_msgs AS MG ON REC.email_msg_id=MG.id  Where REC.id ='" & id & "'  ORDER BY REC.id DESC "

    Set get_email_conf_from_recipient = database.cn.Execute(sqlStrQuery)



End Function



Public Function Create(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    Create = database.Insert("mail_config_msgs", data)

End Function


'/*
'
'
'
'*/
Public Function get_layout(ByVal name As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " SELECT * FROM  app_email_layouts  Where name ='" & name & "'   ORDER BY id DESC LIMIT 1 "

    Set get_layout = database.cn.Execute(sqlStrQuery)



End Function


