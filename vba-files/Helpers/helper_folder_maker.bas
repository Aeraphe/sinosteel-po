Attribute VB_Name = "helper_folder_maker"

'namespace=vba-files\Helpers



Const SpecialCharacters As String = "!,$,%,^,*,{,[,],},?"  'modify as needed

Function get_eng_doc_folder(ByVal project_id As String, ByVal document_id As String, Optional doc_flow_type As String = "COMMENTS") As String

    Dim folders As Object
    Dim full_document_path As String


    Set folders = get_app_folders(doc_flow_type, document_id, project_id)
    'Check if all projects folders exists if not create each one
    Call create_project_root_folder(folders("ROOT_FOLDER"))

    'Add CDOC folder to project folder
    full_document_path = folders("ROOT_FOLDER")

    full_document_path = create_folder_handler(full_document_path)

    'Add Doc category folder
    full_document_path = Trim(full_document_path & "\" & clear_folder_name(folders("DOC_TYPE")) & " - " & clear_folder_name(folders("DOC_CATEGORY")))
    full_document_path = create_folder_handler(full_document_path)

    If (folders("DISCIPLINE") <> "") Then
        full_document_path = full_document_path & "\" & clear_folder_name(folders("DISCIPLINE"))
        full_document_path = create_folder_handler(full_document_path)
    End If

    If (folders("FOLDER") <> "") Then
        full_document_path = full_document_path & "\" & Trim(folders("FOLDER"))
        full_document_path = create_folder_handler(full_document_path)
    End If

    If (folders("CONTRACT_ITEM") <> "") Then
        full_document_path = full_document_path & "\" & clear_folder_name(folders("CONTRACT_ITEM"))
        full_document_path = create_folder_handler(full_document_path)
    End If

    get_eng_doc_folder = full_document_path

End Function



Private Function get_app_folders(ByVal doc_flow_type As String, ByVal document_id As String, ByVal project_id As String) As Object


    Dim app_folders As Object
    Dim project_root_folder As String
    Dim docQuery As ADODB.Recordset
    Dim doc_data As Object
    Dim response As Object
    Dim full_document_path As String

    Set response = CreateObject("Scripting.Dictionary")

    Set app_folders = CreateObject("Scripting.Dictionary")
    Set app_folders = helper_app.get_projec_folders(project_id)


    Set doc_data = CreateObject("Scripting.Dictionary")
    doc_data("PROP1") = document_id

    If (doc_flow_type = "COMMENTS") Then

        project_root_folder = app_folders("ENG_DOC_COMMENT_FULL_PATH")

        Set docQuery = db_documents.get_document_by_review_id(doc_data)

    ElseIf (doc_flow_type = "SENT") Then

        project_root_folder = app_folders("ENG_DOC_SENT_FULL_PATH")

        Set docQuery = db_documents.get_document_by_id(doc_data)
      ElseIf (doc_flow_type = "COMMENTS_WITH_DOC_ID") Then
       project_root_folder = app_folders("ENG_DOC_COMMENT_FULL_PATH")
       Set docQuery = db_documents.get_document_by_id(doc_data)

    End If

    response("DOC_TYPE") = Trim(UCase(XdbFactory.getData(docQuery, "doc_type_code")))
    response("CONTRACT_ITEM") = Trim(UCase(XdbFactory.getData(docQuery, "contract_item")))
    response("DOC_CATEGORY") = Trim(UCase(XdbFactory.getData(docQuery, "category")))
    response("DISCIPLINE") = Trim(UCase(XdbFactory.getData(docQuery, "discipline")))
    response("FOLDER") = Trim(UCase(XdbFactory.getData(docQuery, "folder")))
    response("ROOT_FOLDER") = project_root_folder

    Set get_app_folders = response
End Function

'/*
'
'
'Check is folder project root folder exist and create
'
'*/
Private Function create_project_root_folder(project_root_folder As String) As String
    Dim folders As Variant
    Dim i As Long

    folders = Split(project_root_folder, "\")

    i = 0
    Dim full_document_path As String
    For Each folder In folders

        If (i <> 0 And i <> 1) Then

            full_document_path = create_folder_handler(full_document_path)

        Else
            full_document_path = full_document_path & folder & "\"
        End If
        i = i + 1
    Next folder

    create_project_root_folder = full_document_path
End Function



Private Function create_folder_handler(folder_path As String) As String
    Dim fso As Object
    Dim prepared_folder_path As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    prepared_folder_path = prepar_string_for_create_folder(folder_path)

    If Not fso.FolderExists(prepared_folder_path) Then
    
    On Error Resume Next
        fso.CreateFolder prepared_folder_path

    End If

    create_folder_handler = prepared_folder_path
End Function

Function prepar_string_for_create_folder(myString As String) As String


    Dim newString As String
    Dim char As Variant

    newString = myString
    For Each char In Split(SpecialCharacters, ",")
        newString = Replace(newString, char, " ")
    Next '

    prepar_string_for_create_folder = helper_string.RemoveLineBreak(newString)
End Function

Function clear_folder_name(myString As String) As String

Dim ReplaceSpecialCharacters As String
ReplaceSpecialCharacters = "!,$,%,^,*,{,[,],},/,\,?"  'modify as needed

    Dim newString As String
    Dim char As Variant

    newString = myString
    For Each char In Split(ReplaceSpecialCharacters, ",")
        newString = Replace(newString, char, " ")
    Next '

    clear_folder_name = helper_string.RemoveLineBreak(newString)
End Function
