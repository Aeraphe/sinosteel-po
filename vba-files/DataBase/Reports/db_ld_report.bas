Attribute VB_Name = "db_ld_report"


'namespace=vba-files\DataBase\Reports



'/*
'
'  data("name") = string as project_id
'
'*/
Public Function generate(ByVal project_id As String) As Variant

    Dim database As Object
    Dim data As Object

    Set data = CreateObject("Scripting.Dictionary")
    data("name") = project_id


    Set database = XdbFactory.Create
    Set generate = database.SelectX("get_project_ld", data)

End Function
