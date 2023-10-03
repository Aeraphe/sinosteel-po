Attribute VB_Name = "db_documents"


'namespace=vba-files\DataBase




Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("documents")

End Function


Public Function get_all_project_documents(ByVal project_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select  * FROM  documents  Where  project_id  = " & project_id & "  ORDER BY id DESC "


    Set get_all_project_documents = database.cn.Execute(sqlStrQuery)


End Function



'/*
'
'data(doc_review_id)
'
'*/
Public Function get_document_by_review_id(data As Object) As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set get_document_by_review_id = database.SelectX("get_document_by_review", data)


End Function




Public Function get_document_by_review_id2(ByVal doc_review_id As String) As Variant

    Dim data As Object
    Dim database As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("doc_review_id") = doc_review_id


    Set database = XdbFactory.Create

    Set get_document_by_review_id2 = database.SelectX("get_document_by_review", data)


End Function

Public Function get_document_by_id(data As Object) As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set get_document_by_id = database.SelectX("get_document", data)


End Function



Public Function getDocumentById(doc_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select DOC.*,CAT.name As category  FROM  documents As DOC  INNER JOIN document_categories As CAT ON CAT.id = DOC.category_id   Where  DOC.id  = " & doc_id & "  ORDER BY id DESC "

    Set getDocumentById = database.cn.Execute(sqlStrQuery)


End Function


Public Function getAllExtensions() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAllExtensions = database.getAll("document_extensions")

End Function


Public Function getAllDocCodes() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAllDocCodes = database.getAll("documents_codes")

End Function


Public Function getAllDocFormats() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAllDocFormats = database.getAll("documents_formats")

End Function

Public Function search(ByVal project_id As String, Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "name") As Variant
    If project_id = "" Then Exit Function

        Dim sqlStrQuery As String
        Dim database As Object
        Set database = XdbFactory.Create

        sqlStrQuery = "Select Cat.name As category, DI.name As discipline, DOC.id, DOC.doc_extension, DOC.doc_format, DOC.obs, " & _
        "DOC.sinosteel_doc_number, DOC.doc_number, DOC.name, DOC.description, DOC.pages, DOC.doc_type_code, " & _
        "DOC.contract_item, SUP.name As supplier, APP_FOLDERS.folder " & _
        "FROM documents DOC " & _
        "LEFT JOIN document_categories Cat ON Cat.id = DOC.category_id " & _
        "LEFT JOIN discipline DI ON DI.id = DOC.discipline_id " & _
        "LEFT JOIN suppliers SUP ON SUP.id = DOC.supplier_id " & _
        "LEFT JOIN app_document_folders APP_FOLDERS ON APP_FOLDERS.id = DOC.app_document_folder_id " & _
        "WHERE DOC.project_id = '" & project_id & "'"

        If searchWord <> "" Then
            sqlStrQuery = sqlStrQuery & " And DOC." & filterType & " LIKE '%" & searchWord & "%' "
        End If

        sqlStrQuery = sqlStrQuery & " ORDER BY DOC.id DESC"




        Set search = database.cn.Execute(sqlStrQuery)

End Function

Public Function SearchLastDocumentReview(ByVal project_id As String, Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "name") As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create


    sqlStrQuery = "Select DOC.id,DR.id As rev_id,Cat.name As category,DR.rev_code As last_rev,DR.grd_status_date,DR.grd_status,DR.status, DR.issue ,DOC.doc_extension,DOC.doc_format,DOC.obs, DOC.sinosteel_doc_number,DOC.doc_number, DOC.name,DOC.description,DOC.pages,DOC.doc_type_code,DSC.name As discipline,DOC.contract_item FROM  documents  As DOC LEFT JOIN discipline As DSC ON DSC.id=DOC.discipline_id LEFT JOIN  document_categories As Cat ON Cat.id=DOC.category_id LEFT JOIN documents_reviews As DR ON DR.doc_id=DOC.id Where DOC.project_id='" & project_id & "'  And  DOC." & filterType & "     LIKE '%" & searchWord & "%'  And DR.rev_code in (Select drr.rev_code from documents_reviews As DRR WHERE DOC.id=DRR.doc_id   ORDER BY DRR.id DESC  LIMIT 1 ) ORDER BY DR.rev_code  ASC"
    Set SearchLastDocumentReview = database.cn.Execute(sqlStrQuery)

End Function


Public Function get_last_review_letter(ByVal project_id As String, ByVal doc_id) As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create

    sqlStrQuery = "Select DOC.id,DR.id As rev_id,Cat.name As category,DR.rev_code As last_rev,DR.grd_status_date,DR.grd_status,DR.status, DR.issue ,DOC.doc_extension,DOC.doc_format,DOC.obs, DOC.sinosteel_doc_number,DOC.doc_number, DOC.name,DOC.description,DOC.pages,DOC.doc_type_code FROM  documents  As DOC  LEFT JOIN  document_categories As Cat ON Cat.id=DOC.category_id LEFT JOIN documents_reviews As DR ON DR.doc_id=DOC.id Where DOC.project_id='" & project_id & "'  And  DOC.id='" & doc_id & "'  ORDER BY DR.rev_code  DESC LIMIT 1"
    Set get_last_review_letter = database.cn.Execute(sqlStrQuery)

End Function

Public Function SearchLimit(ByVal project_id As String, Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "name", Optional ByVal limit As String = "1") As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create


    sqlStrQuery = "Select Cat.name As category,DR.rev_code As last_rev,DR.grd_status_date,DR.grd_status,DR.status,DR.id As rev_id, DR.issue ,DOC.id,DOC.doc_extension,DOC.doc_format,DOC.obs, DOC.sinosteel_doc_number,DOC.doc_number, DOC.name,DOC.description,DOC.pages,DOC.doc_type_code,DOC.contract_item,DOC.discipline_id FROM  documents  As DOC  LEFT JOIN  document_categories As Cat ON Cat.id=DOC.category_id  LEFT JOIN documents_reviews As DR ON DR.doc_id=DOC.id Where DOC.project_id=" & project_id & " And  DOC." & filterType & "   LIKE '%" & searchWord & "%'  ORDER BY DR.id DESC LIMIT   " & limit
    Debug.Print sqlStrQuery
    Set SearchLimit = database.cn.Execute(sqlStrQuery)

End Function





Public Function SearchInColumn(ByVal project_id As String, ByVal searchWord As String, ByVal column As String) As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create
    sqlStrQuery = "Select  * FROM  documents  Where  project_id=" & project_id & " And " & column & " = '" & searchWord & "'  ORDER BY id DESC  LIMIT 1"
    Set SearchInColumn = database.cn.Execute(sqlStrQuery)

End Function



Public Function SearchLastRev(ByVal doc_id As String, Optional ByVal limit As Long = 1) As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create
    sqlStrQuery = "Select  * FROM  documents_reviews  Where  doc_id=" & doc_id & "  ORDER BY id DESC  LIMIT " & limit
    Set SearchLastRev = database.cn.Execute(sqlStrQuery)

End Function


Public Function SearchReviews(ByVal doc_id As String) As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create
    sqlStrQuery = "Select  * FROM  documents_reviews  Where  doc_id=" & doc_id & "    ORDER BY id DESC "
    Set SearchReviews = database.cn.Execute(sqlStrQuery)

End Function

Public Function getDocIssueTypes() As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create

    Set getDocIssueTypes = database.getAll("documents_issue_types")

End Function


Public Function getDocStatusTypes() As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create

    Set getDocStatusTypes = database.getAll("document_status_types")

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

    sqlStrQuery = " DELETE   FROM  documents  Where id = " & idData

    database.cn.Execute (sqlStrQuery)



End Function




Public Function Import(ByRef document As Variant) As Long


    Dim database As Object
    Dim doc_id As Long

    Set database = XdbFactory.Create

    If Not document Is Nothing Then
        Import = database.Insert("documents", document)
    Else
        Import = 0
    End If



End Function

'/*
'
'Create Sppliers
'
'@param <Array>  data
'
'*/
Public Function Create(ByRef document As Variant, ByRef document_properties As Variant, ByRef document_equipaments As Variant) As Long


    Dim database As Object
    Dim doc_id As Long

    Set database = XdbFactory.Create

    If Not document Is Nothing Then
        doc_id = database.Insert("documents", document)
        Call InsertDocumentPropertiesHandler(document_properties, doc_id)
        Call InsertDocumentEquipamentsHandler(document_equipaments, doc_id)

    End If




End Function


Function InsertDocumentPropertiesHandler(ByRef properties_list As Variant, doc_id As Long)


    Dim database As Object
    Set database = XdbFactory.Create

    Dim i As Long

    For i = 0 To properties_list.ListCount - 1
        Dim items As Object
        Set items = CreateObject("Scripting.Dictionary")

        items("document_id") = doc_id
        items("property_id") = properties_list.List(i, 0)
        items("value") = UCase(properties_list.List(i, 2))

        Call database.Insert("document_properties", items)

    Next i

End Function

Function insert_property(ByRef Property As Variant) As Long
    Dim database As Object
    Dim doc_id As Long

    Set database = XdbFactory.Create

    If Not Property Is Nothing Then
        insert_property = database.Insert("documents", Property)

    End If

    insert_property = 0

End Function


Private Function InsertDocumentEquipamentsHandler(ByRef equipaments_list As Variant, doc_id As Long)


    Dim database As Object
    Set database = XdbFactory.Create
    Dim i As Long

    For i = 0 To equipaments_list.ListCount - 1
        Dim items As Object
        Set items = CreateObject("Scripting.Dictionary")

        items("document_id") = doc_id
        items("equipament_id") = equipaments_list.List(i, 0)
        Call database.Insert("documents_equipaments", items)

    Next i

End Function


'/*
'
'Create
'
'@param <Array>  data
'
'*/
Public Function InsertDocumentReview(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    InsertDocumentReview = database.Insert("documents_reviews", data)

End Function




Public Function getDocumentReviews(doc_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select  * FROM  documents_reviews  Where  doc_id  = " & doc_id & "  ORDER BY id DESC "


    Set getDocumentReviews = database.cn.Execute(sqlStrQuery)


End Function


Public Function get_doc_review_issue(doc_id As String, rev_code As String, issue As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select  * FROM  documents_reviews  Where  doc_id  = '" & doc_id & "' And rev_code='" & rev_code & "' And issue='" & issue & "'"


    Set get_doc_review_issue = database.cn.Execute(sqlStrQuery)


End Function

Public Function search_doc_by_review(doc_id As String, rev_code As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select  * FROM  documents_reviews  Where  doc_id  = '" & doc_id & "' And rev_code='" & rev_code & "'"


    Set search_doc_by_review = database.cn.Execute(sqlStrQuery)


End Function

Public Function get_first_review(doc_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select  * FROM  documents_reviews  Where  doc_id  = '" & doc_id & "' order by rev_code  desc  Limit 1"


    Set get_first_review = database.cn.Execute(sqlStrQuery)


End Function


Public Function get_last_review(doc_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select  * FROM  documents_reviews  Where  doc_id  = '" & doc_id & "' order by rev_code   Limit 1"


    Set get_last_review = database.cn.Execute(sqlStrQuery)


End Function



Public Function updateStatus(ByVal data As Variant, ByVal where As String)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Call database.update("documents_reviews", data, where)

End Function



Public Function update(ByVal data As Variant, ByVal where As String)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Call database.update("documents", data, where)

End Function


'vba Function For excel
Public Function get_doc_by_review2(ByVal review_id As String) As Variant
    Dim sqlStrQuery As String
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim database  As Object
    Set database = XdbFactory.Create
    Set conn = database.cn

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select DOC.project_id,DOC.id, DOC.doc_number, DOC.doc_extension, DR.rev_code, DOC.name, DOC.description, DR.request_doc_id FROM documents_reviews As DR INNER JOIN documents As DOC ON DOC.id = DR.doc_id WHERE DR.id = ?"
    cmd.Parameters.Append cmd.CreateParameter("review_id", adVarChar, adParamInput, 50, review_id)
    cmd.ActiveConnection = conn

    Set rs = cmd.Execute
    Set get_doc_by_review2 = rs
End Function

Public Function get_doc_by_review(ByVal review_id As String) As Variant
    Dim sqlStrQuery As String
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim database  As Object
    Set database = XdbFactory.Create
    Set conn = database.cn

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = "Select DOC.project_id,DOC.id, DOC.doc_number, DOC.doc_extension, DR.rev_code, DOC.name, DOC.description, DR.request_doc_id FROM documents_reviews As DR INNER JOIN documents As DOC ON DOC.id = DR.doc_id WHERE DR.id = ?"
    cmd.Parameters.Append cmd.CreateParameter("review_id", adVarChar, adParamInput, 50, review_id)
    cmd.ActiveConnection = conn

    Set rs = cmd.Execute
    get_doc_by_review = rs.GetRows

    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Function





'/*
'
'Create
'
'@param <Array>  data
'
'*/
Public Function add_document_rejected(ByRef data As Object) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    add_document_rejected = database.Insert("document_reviews_reject", data)

End Function


Public Function update_reject_document_state(ByRef data As Variant, ByVal where As String)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Call update_reject_document_state.update("document_reviews_reject", data, where)

End Function


Public Function update_review_issue(ByVal review_id As String, ByVal issue As String)


    Dim data As Object
    Dim sqlStrQuery As String
    Dim database As Object
    Dim where As String

    Set data = CreateObject("Scripting.Dictionary")

    data("issue") = issue


    Set database = XdbFactory.Create

    where = "id='" & review_id & "'"

    Call database.update("documents_reviews", data, where)

End Function


Public Function update_review_status(ByVal review_id As String, ByVal status As String, ByVal status_date As String)


    Dim data As Object
    Dim sqlStrQuery As String
    Dim database As Object
    Dim where As String

    Set data = CreateObject("Scripting.Dictionary")

    data("status") = status
    data("status_date") = DateHelpers.FormatDateToSQlite(status_date)

    Set database = XdbFactory.Create

    where = "id='" & review_id & "'"

    Call database.update("documents_reviews", data, where)

End Function
