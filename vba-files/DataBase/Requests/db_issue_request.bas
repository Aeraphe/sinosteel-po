Attribute VB_Name = "db_issue_request"



'namespace=vba-files\DataBase\Requests






Public Function get_request(ByVal request_id As String) As Variant

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT RQ.title,RQ.msg,RQ.category,RQ.folder,RQ.status,RQ.create_ate,PR.name,PR.project_code,(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='EMITIR')as docs_to_post FROM  eng_request_issue AS RQ  INNER JOIN projects AS PR ON PR.id = RQ.project_id Where  RQ.id  = '" & request_id & "' ORDER BY RQ.id DESC LIMIT 1"
    Debug.Print sqlStrQuery
    Set get_request = database.cn.Execute(sqlStrQuery)

End Function







Public Function getDocRequestData(ByVal docRequestId As String) As Variant


    Dim data As Object
    Dim database As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("docRequestId") = docRequestId


    Set database = XdbFactory.Create

    Set getDocRequestData = database.SelectX("GET_DOC_RESQUEST_DATA", data)


End Function


Public Function get_request_docs(ByVal request_id As String) As Variant


    Dim data As Object
    Dim database As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("requestId") = request_id


    Set database = XdbFactory.Create

    Set get_request_docs = database.SelectX("GET_ENG_REQUEST_DOCS", data)


End Function



Public Function getAllRequestsFromProject(ByVal projectId As String) As Variant

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT US.name as user_name,RQ.id,RQ.title,RQ.msg,RQ.category,RQ.folder,RQ.status,RQ.create_ate,PR.name,PR.project_code," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='PEND. EXT.')as docs_pend_ext," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='NO FLUXO')as total_docs_in_flow," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='CONCLUIDO')as docs_post," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='ENVIADO')as docs_sent," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='REJEITADO')as docs_rejected," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='PROGRAMADO')as docs_schedule," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='HOLD')as docs_hold," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='CANCELADO')as docs_canceled," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='EMITIR')as docs_to_send," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id AND ERR.status='LIB. ENG')as docs_lib_eng," & _
    "(SELECT COUNT(*) FROM eng_request_issue_documents AS ERR WHERE ERR.eng_request_issue_id=RQ.id)as total_docs  FROM  eng_request_issue AS RQ  INNER JOIN projects AS PR ON PR.id = RQ.project_id INNER JOIN users AS US ON US.id=RQ.user_id Where  PR.id  = '" & projectId & "' ORDER By RQ.id DESC"
    Debug.Print sqlStrQuery
    Set getAllRequestsFromProject = database.cn.Execute(sqlStrQuery)

End Function
'/*
'
'
' data("status") = doc_review_id
' data("status_date") = doc_review_id
' data("user_id_doc_flow") = Auth.get_user_i
'
'*/
Public Function updateRequestDocument(ByVal data As Variant, ByVal docRequestId As String)

    Dim sqlStrQuery As String

    Dim database As Object
    Set database = XdbFactory.Create

    Dim where As String

    where = " id='" & docRequestId & "'"

    Call database.update("eng_request_issue_documents", data, where)

End Function


Public Function getRequestDocsSentToFactory(ByVal request_id As String) As Variant


    Dim data As Object
    Dim database As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("requestId") = request_id


    Set database = XdbFactory.Create

    Set getRequestDocsSentToFactory = database.SelectX("GET_DOCS_SENT_TO_FACTORY", data)


End Function



'/*
'
'This function it will be used for fixed documents that was post without requests
'
'*/
Public Function fixDocRequestId(ByVal requestId As String, ByVal docId As String, ByVal docRevCode As String)



    Dim database As Object
    Set database = XdbFactory.Create

    Set data = CreateObject("Scripting.Dictionary")
    data("request_doc_id") = requestId

    Dim where As String

    where = " doc_id='" & docId & "' AND rev_code='" & docRevCode & "'"

    Call database.update("documents_reviews", data, where)

End Function


Public Function getRequestDocsWithProps(ByVal request_id As String) As Variant


    Dim data As Object
    Dim database As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("requestId") = request_id


    Set database = XdbFactory.Create

    Set getRequestDocsWithProps = database.SelectX("GET_REQ_DOCS_WITH_PROPS", data)


End Function


Public Function searchRequestedDocument(ByVal SearchText As String, ByVal category As String, Optional ByVal status As String) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim database As Object

    Set database = XdbFactory.Create
    Set conn = database.cn

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT RDOC.eng_request_issue_id AS red_number, " & _
                      "DOC.id, DOC.doc_number, DOC.name, DOC.description, " & _
                      "RDOC.priority, RDOC.id As doc_request_id, RDOC.rev_code As rev_code_request, " & _
                      "RDOC.file_name, RDOC.file_path, RDOC.issue, RDOC.file_folder, " & _
                      "RDOC.file_extension, RDOC.status, RDOC.status_date, RDOC.post_in_date, " & _
                      "RDOC.post_user_response_msg, DOC.contract_item, REV.id As rev_id, " & _
                      "REV.status As rev_status, US.name As response_user " & _
                      "FROM eng_request_issue_documents As RDOC " & _
                      "LEFT JOIN USERS As US ON US.id = RDOC.post_user_response_id " & _
                      "INNER JOIN eng_request_issue As ENG_REQ ON ENG_REQ.id = RDOC.eng_request_issue_id " & _
                      "INNER JOIN documents_full_search As DOC_FS ON DOC_FS.search_doc_id = RDOC.doc_id " & _
                      "INNER JOIN documents As DOC ON DOC.id = RDOC.doc_id " & _
                      "LEFT JOIN projects As PRJ ON PRJ.id = DOC.project_id " & _
                      "LEFT JOIN documents_reviews As REV ON REV.request_doc_id = RDOC.id " & _
                      "WHERE DOC_FS.search_content LIKE ? " & _
                      "AND RDOC.status NOT IN ('CONCLUIDO', 'REJEITADO', 'CANCELADO') " & _
                      "AND ENG_REQ.category = ? "

    If (status <> "") Then
       cmd.CommandText = cmd.CommandText & " AND RDOC.status=" & "'" & status & "' "
    End If
    
    cmd.CommandText = cmd.CommandText & " ORDER BY RDOC.eng_request_issue_id DESC"

    cmd.Parameters.Append cmd.CreateParameter("searchText", adVarChar, adParamInput, 50, "%" & SearchText & "%")
    cmd.Parameters.Append cmd.CreateParameter("category", adVarChar, adParamInput, 50, category)

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd

    Set searchRequestedDocument = rs
    



End Function



