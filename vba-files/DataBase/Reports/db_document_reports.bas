Attribute VB_Name = "db_document_reports"


'namespace=vba-files\DataBase\Reports



'/*
'
'  data("name") = string as project_id
'
'*/
Public Function get_documents_today_status_change(data As Variant) As Variant

    Dim database As Object

    Set database = XdbFactory.Create
    Set get_documents_today_status_change = database.SelectX("report_doc_satus_change_today", data)

End Function
