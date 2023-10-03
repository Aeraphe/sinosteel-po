Attribute VB_Name = "db_grd"


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
    Create = database.Insert("grd_recipients", data)

End Function



Public Function CreateGRD(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    CreateGRD = database.Insert("grd", data)

End Function


Public Function insertGRDDocuments(ByRef data As Object) As Long

    Dim sqlStrQuery As String
    Dim rs As Object
    Dim database As Object
    Dim existentRecord As Boolean

    Set database = XdbFactory.Create

    ' Check If the record already exists
    sqlStrQuery = "Select soft_delete FROM grd_documents WHERE grd_id=" & data("grd_id") & " And doc_rev_id=" & data("doc_rev_id")
    Set rs = database.cn.Execute(sqlStrQuery)

    existentRecord = Not (rs.EOF And rs.BOF)

    If existentRecord Then
        ' If the record exists And soft_delete is 1, update it
        If rs.fields("soft_delete").Value = 1 Then
            sqlStrQuery = "UPDATE grd_documents Set soft_delete=0 WHERE grd_id=" & data("grd_id") & " And doc_rev_id=" & data("doc_rev_id")
            database.cn.Execute (sqlStrQuery)
        End If
    Else
        ' If the record doesn't exist, insert it
        insertGRDDocuments = database.Insert("grd_documents", data)
    End If

End Function


Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("grd_recipients")

End Function


Public Function getMediaTypes() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getMediaTypes = database.getAll("grd_media_types")

End Function

Public Function getContentTypes() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getContentTypes = database.getAll("grd_content_types")

End Function

'/*
'
'Delete grd By ID
'
'*/
Public Function delete(ByRef idData As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  grd  Where id = " & idData

    database.cn.Execute (sqlStrQuery)



End Function

'/*
'
'Delete grd By ID
'
'*/
Public Function delete_document(ByRef grd_id As String, doc_rev_id As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  grd_documents  Where grd_id = " & grd_id & " And doc_rev_id = " & doc_rev_id

    database.cn.Execute (sqlStrQuery)



End Function



'/*
'
'Get last Recipent
'
'*/
Public Function getLastGRDFromRecipient(ByRef recipent_id As String) As Variant

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " Select *   FROM  grd  Where recipent_id = " & recipent_id & " ORDER BY id DESC LIMIT 1"

    Set getLastGRDFromRecipient = database.cn.Execute(sqlStrQuery)



End Function



'Delete budget By ID
'
'*/
Public Function getAllGRDFromRecipient(ByRef recipent_id As String) As Variant

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " Select GRD.id,GRD.user_id,GRD.sequece_number,GRD.description,GRD.issue_date,GRD.confirmation_date,GRD.obs,GRD.create_ate  FROM  grd   Where recipent_id = " & recipent_id & " And soft_delete='0' ORDER BY id DESC "

    Set getAllGRDFromRecipient = database.cn.Execute(sqlStrQuery)



End Function


Public Function getGRDItems(data As Variant) As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getGRDItems = database.SelectX("get_grd_items", data)

End Function


Public Function get_grd_sent_to_project_contractor(data As Variant) As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set get_grd_sent_to_project_contractor = database.SelectX("get_grd_sent_to_project_contractor", data)

End Function



Public Function getById(id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "Select GRD.recipent_id, grd.sequece_number,grd.description,grd.issue_date,RE.code,RE.folder_name,RE.project_id,RE.email_msg_id,SUP.name,SUP.code As sup_code,SUP.email,SUP.person,GRD.user_id FROM  grd INNER JOIN grd_recipients As RE ON grd.recipent_id=RE.id  INNER JOIN suppliers As SUP ON SUP.id=RE.supplier_id  Where  grd.id  = " & id & "  ORDER BY grd.id DESC "


    Set getById = database.cn.Execute(sqlStrQuery)


End Function


Public Function create_sandbox(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    create_sandbox = database.Insert("grd_sandbox", data)

End Function


Public Function insert_doc_to_sandbox(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    insert_doc_to_sandbox = database.Insert("grd_sandbox_items", data)

End Function



Public Function update(ByVal data As Variant, ByVal where As String)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Call database.update("grd", data, where)

End Function

'/*
' This Function soft deletes a document from the database based on the grd_id And doc_rev_id parameters.
' The Function takes two string parameters: grdId And docReviewId.
'*/
Public Function softDeleteDocument(ByVal grdId As String, ByVal docReviewId As String) As Boolean

    ' Declaring all local variables.
    Dim sqlStrQuery As String
    Dim database As Object
    Dim cmd As New ADODB.Command

    sqlStrQuery = "UPDATE grd_documents Set soft_delete = 1 WHERE grd_id = ? And doc_rev_id = ?"

    Set database = XdbFactory.Create

    cmd.ActiveConnection = database.cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sqlStrQuery

    ' Define parameters
    cmd.Parameters.Append cmd.CreateParameter("grd_id", adInteger, adParamInput, 255, grdId)
    cmd.Parameters.Append cmd.CreateParameter("doc_rev_id", adInteger, adParamInput, , docReviewId)

    On Error GoTo CatchError
        cmd.Execute

        ' Return True If there were no errors:
        softDeleteDocument = True
     Exit Function

CatchError:
        helper_log.DebugApp "Try To softDeleteDocument a GRD Document And an error occurred: " & Err.description
        softDeleteDocument = False
End Function





Public Function update_document(ByVal data As Variant, ByVal where As String)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Call database.update("grd_documents", data, where)

End Function



'/*
'
'
'*/
Public Function get_who_received_the_documents(ByVal project_id As String, ByVal review_id As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("review_id") = review_id
    data("project_id") = project_id

    Dim database As Object
    Set database = XdbFactory.Create

    Set get_who_received_the_documents = database.SelectX("get_who_received_documents", data)

End Function




'/*
'
'
'*/
Public Function getDocumentsNotReturnFromApproveFlow(ByVal project_id As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")


    data("project_id") = project_id

    Dim database As Object
    Set database = XdbFactory.Create

    Set getDocumentsNotReturnFromApproveFlow = database.SelectX("DOCS_NOT_RETURN_FROM_CONTRACTOR", data)

End Function


'/*
'
'
'*/
Public Function gel_all_doc_replaced_pend_to_submitting(ByVal project_id As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")


    data("project_id") = project_id

    Dim database As Object
    Set database = XdbFactory.Create

    Set gel_all_doc_replaced_pend_to_submitting = database.SelectX("get_all_doc_replaced_pend_to_submitting", data)

End Function


'/*
'
'
'*/
Public Function get_documents_not_sent_to_contractor(ByVal project_id As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")


    data("project_id") = project_id

    Dim database As Object
    Set database = XdbFactory.Create

    Set get_documents_not_sent_to_contractor = database.SelectX("get_documents_not_sent_to_contractor", data)

End Function
'/*
'
'
'*/
Public Function get_doc_from_contractor_grd_recipient(ByVal project_id As String, review_id As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("project_id") = project_id
    data("review_id") = review_id

    Dim database As Object
    Set database = XdbFactory.Create

    Set get_doc_from_contractor_grd_recipient = database.SelectX("get_doc_from_contractor_grd_recipient", data)

End Function


Public Function check_if_has_last_doc_review(ByVal project_id As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("project_id") = project_id


    Dim database As Object
    Set database = XdbFactory.Create

    Set check_if_has_last_doc_review = database.SelectX("check_if_has_last_doc_review", data)

End Function




Public Function getGrdFromRequestDoc(ByVal requestDocId As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("requestDocId") = requestDocId


    Dim database As Object
    Set database = XdbFactory.Create

    Set getGrdFromRequestDoc = database.SelectX("GET_GRDS_FROM_REQUEST_DOCS", data)

End Function




Public Function getContractorGrdByDocReviewId(ByVal docReviewId As String) As ADODB.Recordset

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("docReviewId") = docReviewId


    Dim database As Object
    Set database = XdbFactory.Create

    Set getContractorGrdByDocReviewId = database.SelectX("GET_CONTRATOR_GRD_FROM_DOC_REVIEW", data)

End Function
