Attribute VB_Name = "helper_app"



'namespace=vba-files\Helpers



'/*
'
'
'Get Project and app folders
'
'@return reponse:Dictionary
'
' ENG_DOC_FOLDER
' ENG_GRD_FULL_PATH
' ENG_REPORTS_FULL_PATH
' ENG_DOC_REPORTS
' ENG_GRD_SENDED
' ENG_GRD_RECEIVE
' ENG_CDOC
' ENG_GRD
' ENG_ROOT_FOLDER
' ENG_ROOT_DOC_FOLDER
' CDOC_ROOT_FOLDER
'
'*/
 Function get_projec_folders(ByVal project_id As String) As Object

    Dim response As Object
    Dim projectQuery As ADODB.Recordset
    Dim appFoldersQuery As ADODB.Recordset
    Dim prop As String
    Dim Value As String
    Dim PROJECT_FOLDER As String
    Dim driver As String

    driver = config_sheet.Range("CONFIG_LAN_PATH").Value
    
    Set projectQuery = db_projects.get_by_id(project_id)
    Set response = CreateObject("Scripting.Dictionary")
    
    Set appFoldersQuery = db_app_folders.get_all()

    Do Until appFoldersQuery.EOF

        Value = helper_string.RemoveLineBreak(appFoldersQuery.fields.item("value"))
        prop = helper_string.RemoveLineBreak(appFoldersQuery.fields.item("prop"))
        response(prop) = Value

        appFoldersQuery.MoveNext
    Loop

    PROJECT_FOLDER = driver & response("ENG_ROOT_FOLDER") & "\" & projectQuery.fields.item("project_folder") & "\" & response("ENG_ROOT_DOC_FOLDER")

    response("ENG_DOC_FOLDER") = PROJECT_FOLDER
    response("ENG_GRD_FULL_PATH") = PROJECT_FOLDER & "\" & response("ENG_CDOC") & "\" & response("ENG_GRD")
    response("ENG_REPORTS_FULL_PATH") = PROJECT_FOLDER & "\" & response("ENG_CDOC") & "\" & response("ENG_DOC_REPORTS")
    response("ENG_DOC_COMMENT_FULL_PATH") = PROJECT_FOLDER & "\" & response("ENG_DOCS_COMMENTS")
    response("ENG_DOC_SENT_FULL_PATH") = PROJECT_FOLDER & "\" & response("ENG_DOCS_SENT")
    response("ENG_DOC_REJECTED_FULL_PATH") = PROJECT_FOLDER & "\" & response("ENG_DOCS_REJECTED")

    Set get_projec_folders = response



End Function


