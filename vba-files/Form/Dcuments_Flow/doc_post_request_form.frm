VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_post_request_form 
   Caption         =   "Requisições De Emissão de Documentos (RED)"
   ClientHeight    =   13830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   26310
   OleObjectBlob   =   "doc_post_request_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_post_request_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Dcuments_Flow


Private project_selected_id As String
Private projectSelectedName As String
Public docsSelectedToPost As Object
Private requestIdSelected As String
Private document_selected_id As String
Public docToPostList As Object


Const REQ_TB_COLN_GRD_ID As Long = 10
Const REQ_TB_COLN_DOC_ID As Long = 11
Const REQ_TB_COLN_DOC_REV_ID As Long = 12
Const REQ_TB_COLN_DOC_FILE_PATH As Long = 9
Const REQ_TB_COLN_STATUS As Long = 4
Const REQ_TB_COLN_DOC_NUMBER As Long = 2





Private Sub UserForm_Initialize()

   loadSelectUsers
   PopulateDocRequestStatusDropdown
   GetDocIssueTypesHandler
   getDocExtensionstHandler
   getDocFormatstHandler
   Call Shared_DocCategorySelectComp.Mount(category_select)
   Set docToPostList = CreateObject("Scripting.Dictionary")
   Set docsSelectedToPost = CreateObject("Scripting.Dictionary")

   select_flow_type.AddItem "FABRICACAO"
   select_flow_type.AddItem "PRINCIPAL"
   select_status_filter.AddItem ""
   select_status_filter.AddItem Constants.EMITIR
   select_status_filter.AddItem Constants.SUBISTITUIR
   select_status_filter.AddItem Constants.NO_FLUXO
   select_status_filter.AddItem Constants.ENVIADO
   select_status_filter.AddItem Constants.LIB_ENG
   select_status_filter.AddItem Constants.PEND
   select_status_filter.AddItem Constants.PROGRAMADO


End Sub






Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)
   lb_user.Caption = "( " & Left(auth.user_name, 15) & " )"

End Sub



Private Function getProjectContractItemsHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_projects.get_contract_items(project_selected_id)

   contract_item_select.Clear
   Call Shared_CommonSelectComp.Mount(contract_item_select, respQuery)
End Function

Private Function getDocExtensionstHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllExtensions()

   extension_select.Clear
   Call Shared_CommonSelectComp.Mount(extension_select, respQuery)
End Function

Private Function getDocFormatstHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllDocFormats()

   doc_format_select.Clear
   Call Shared_CommonSelectComp.Mount(doc_format_select, respQuery)
End Function

Private Function GetDocIssueTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getDocIssueTypes()
   issue_select.Clear
   Do Until respQuery.EOF

      issue_select.AddItem XdbFactory.getData(respQuery, "tag")
      issue_select.List(issue_select.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

      respQuery.MoveNext

   Loop

End Function

Private Sub ck_show_item_canceled_Click()
   Call getDocumentsBySelectedResquest(doc_list.List(doc_list.ListIndex, 0))
End Sub

Private Sub ck_show_item_post_Click()
   Call getDocumentsBySelectedResquest(doc_list.List(doc_list.ListIndex, 0))
End Sub

Private Sub ck_show_item_rejected_Click()
   Call getDocumentsBySelectedResquest(doc_list.List(doc_list.ListIndex, 0))
End Sub



Private Sub ck_check_factory_Click()
   If (ck_check_factory.Value) Then
      MultiPage1.Pages("Page2").Enabled = ck_check_factory.Value
      MultiPage1.Pages("Page2").Visible = ck_check_factory.Value
      MultiPage1.Pages("Page1").Enabled = Not ck_check_factory.Value
      MultiPage1.Pages("Page1").Visible = Not ck_check_factory.Value

      MultiPage1.Value = 1

   Else
      MultiPage1.Pages("Page2").Enabled = ck_check_factory.Value
      MultiPage1.Pages("Page2").Visible = ck_check_factory.Value
      MultiPage1.Pages("Page1").Enabled = Not ck_check_factory.Value
      MultiPage1.Pages("Page1").Visible = Not ck_check_factory.Value
      MultiPage1.Value = 0

   End If

End Sub

Private Sub ck_factory_display_Click()
   If (project_txt.Value <> "") Then
      Call getRequestsHandler
   End If
End Sub

Private Sub ck_show_completed_Click()
   If (project_txt.Value <> "") Then
      Call getRequestsHandler
   End If
End Sub



Private Sub ck_show_prin_Click()
   If (project_txt.Value <> "") Then
      Call getRequestsHandler
   End If
End Sub


Private Sub request_items_list_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   Dim doc_file As String

   doc_file = request_items_list.List(request_items_list.ListIndex, 9)

   If (doc_file <> "") Then

      temp_file_path = action_file.copy_to_temp_folder(doc_file)
      Call file_helper.open_file(temp_file_path)

   End If
End Sub


Private Function PopulateDocRequestStatusDropdown()
   Dim requestStatuses As Variant
   requestStatuses = Array(Constants.EMITIR, Constants.NO_FLUXO, Constants.ENVIADO, Constants.CONCLUIDO, Constants.PEND, Constants.CANCELADO, Constants.REJEITADO, Constants.PROGRAMADO)

   Dim i As Integer
   For i = LBound(requestStatuses) To UBound(requestStatuses)
      select_doc_request_status.AddItem requestStatuses(i)
   Next i
End Function

Private Function loadSelectUsers()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_users.getUsersBySector("CDOC")

   user_select.Clear
   Call Shared_CommonSelectComp.Mount(user_select, respQuery)
End Function





'/*
' This Sub-routine is used To handle a button click event For changing the status of a document request. It takes no arguments.
'*/
Private Sub btn_change_doc_status_Click()

   Dim newStatus As String
   newStatus = select_doc_request_status.Value
   If (act_doc_change_status.ConfirmDocumentStatus(newStatus)) Then
      Call ChangeDocumentRequestStatusHandler(newStatus)
   End If

End Sub


'/*
'
' This Function is used To handle the change in status of a document request. It takes no arguments.
'
'*/
Private Function ChangeDocumentRequestStatusHandler(ByVal newStatus As String)

   Dim currentStatus As String

   Dim documentSelected As String
   Dim responseStatusChange As String
   Dim isStatusChange As Object
   Dim grdId As String
   Dim docReviewId As String
   Dim payload As Object




   If newStatus = "" Then
      MsgBox "Error: Selecione um novo Status", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!"
    Exit Function
   End If



   For i = 0 To request_items_list.ListCount - 1
      If request_items_list.Selected(i) Then
         documentSelected = request_items_list.List(i, REQ_TB_COLN_DOC_NUMBER)
         currentStatus = request_items_list.List(i, REQ_TB_COLN_STATUS)
         grdId = request_items_list.List(i, REQ_TB_COLN_GRD_ID)
         docReviewId = request_items_list.List(i, REQ_TB_COLN_DOC_REV_ID)

         Set payload = BuildPayload(grdId, docReviewId, documentSelected)
         Set isStatusChange = act_doc_change_status.ChangeStatus(currentStatus, newStatus, request_items_list.List(i, 13), payload)
         responseStatusChange = responseStatusChange & documentSelected & " : " & isStatusChange("INFO") & vbNewLine
         itemsSelected = True
      End If
   Next i

   If Not itemsSelected Then
      MsgBox "Error: Nenhum Documento Foi Selecionado", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!"
    Exit Function
   End If

   getDocsFromSelectedRequestHandler

   ' Show message box With responseStatusChange As the message text
   sh_edit_text_form.msg_txt = responseStatusChange
   sh_edit_text_form.Show

   Call helper_log.createLogFile(vbNewLine & responseStatusChange, CHANGE_REQUEST_DOC_STATUS_FILE_NAME)

End Function




'/*
' Function To build payload For status change request based on data entered in the form
'*/
Private Function BuildPayload(ByVal grdId As String, ByVal docReviewId As String, ByVal documentSelected As String) As Object
   Dim payload As Object
   Set payload = CreateObject("Scripting.Dictionary")

   payload("INFO") = post_msg_txt.Value
   payload("SCHEDULE_DATE") = post_data_txt.Value
   payload("USER_OWNER_ID") = user_select.Value
   payload("GRD_ID") = grdId
   payload("DOC_REVIEW_ID") = docReviewId
   payload("DOC_NUMBER") = documentSelected

   Set BuildPayload = payload
End Function





Private Sub doc_list_Change()
   grd_list.Clear
   props_list.Clear
   clear_form_fields
   getDocsFromSelectedRequestHandler

End Sub




Private Function getDocsFromSelectedRequestHandler()
   Dim requestId As String

   For i = 0 To doc_list.ListCount - 1
      On Error Resume Next
      If doc_list.Selected(i) = True Then
         requestIdSelected = doc_list.List(i, 0)
         Call getDocumentsBySelectedResquest(requestIdSelected)
         Call checkDocsInFactory(requestIdSelected)
       Exit Function
      End If
   Next i
End Function


Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")


   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      projectSelectedName = project_txt.Value
      project_selected_id = data("id")
      Call getRequestsHandler
      getProjectContractItemsHandler
   Else



   End If
End Sub

Private Sub btn_reload_project_request_search_Click()
   If (project_txt.Value <> "") Then
      Call getRequestsHandler
   End If
End Sub


Private Function getRequestsHandler()

   Dim userLoedgName As String
   Dim userName As String
   userLogedName = auth.user_name

   doc_list.Clear
   request_items_list.Clear

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_issue_request.getAllRequestsFromProject(project_selected_id)

   Dim folderPath As String


   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")


   Dim fObj As Object

   Dim countDocsToPost As Long
   Dim countDocsSent As Long
   Dim countDocsRejected As Long
   Dim countDocsInTheFlow As Long
   Dim countDocs As Long
   Dim countDocPentExt As Long
   Dim scheduleDocs As Long
   Dim countReqPend As Long
   Dim countDocsToSend  As Long
   Dim countDocLibEng As Long

   Dim totalRequestDocsPend As Long
   Dim totalRequestDocsSent As Long
   Dim totalRequestDocsPost As Long
   Dim totalRequestDocsRejected As Long
   Dim totalRequestDocs As Long
   Dim totalScheduleDocs As Long
   Dim countDocCanceled As Long
   Dim totalCanceledDocs As Long
   Dim totalDocsInTheFlow As Long
   Dim totalDocsToSend  As Long
   Dim totalDocLibEng As Long
   Dim totalDocsPentExt As Long

   Dim totalRequestDocsUserPend As Long


   Dim totalDocsPend As Long

   Dim docsPend As Long

   Dim totalFactoryFlowDocsPend As Long
   Dim totalPrincipalFlowDocsPend As Long

   Dim countReqCompleted As Long
   countReqPend = 0



   Dim printDoc As Boolean
   printDoc = False

   Dim flowCategory As String



   Do Until respQuery.EOF

      flowCategory = XdbFactory.getData(respQuery, "category")


      userName = XdbFactory.getData(respQuery, "user_name")
      countDocsPost = XdbFactory.getData(respQuery, "docs_post")
      countDocsSent = XdbFactory.getData(respQuery, "docs_sent")
      countDocsRejected = XdbFactory.getData(respQuery, "docs_rejected")
      countDocs = XdbFactory.getData(respQuery, "total_docs")
      countDocsInTheFlow = XdbFactory.getData(respQuery, "total_docs_in_flow")
      countDocPentExt = XdbFactory.getData(respQuery, "docs_pend_ext")
      countDocCanceled = XdbFactory.getData(respQuery, "docs_canceled")
      countScheduleDocs = XdbFactory.getData(respQuery, "docs_schedule")
      countDocsToSend = XdbFactory.getData(respQuery, "docs_to_send")
      countDocLibEng = XdbFactory.getData(respQuery, "docs_lib_eng")

      docsPend = (countDocs - countDocsPost - countDocsSent - countDocsRejected - countDocCanceled)


      totalScheduleDocs = totalScheduleDocs + countScheduleDocs
      totalRequestDocsSent = totalRequestDocsSent + countDocsSent
      totalRequestDocs = totalRequestDocs + countDocs
      totalRequestDocsPost = totalRequestDocsPost + countDocsPost
      totalRequestDocsRejected = totalRequestDocsRejected + countDocsRejected
      totalCanceledDocs = totalCanceledDocs + countDocCanceled
      totalDocsPend = totalDocsPend + docsPend
      totalDocsInTheFlow = totalDocsInTheFlow + countDocsInTheFlow
      totalDocsToSend = totalDocsToSend + countDocsToSend
      totalDocLibEng = totalDocLibEng + countDocLibEng
      totalDocsPentExt = totalDocsPentExt + countDocPentExt






      If (countDocsPost + countDocsRejected + countDocCanceled = countDocs) Then
         status = "FINALIZADO"
         countReqCompleted = countReqCompleted + 1
      Else
         status = "PENDENTE"
         countReqPend = countReqPend + 1
      End If

      If (userLogedName = userName) Then
         totalRequestDocsUserPend = totalRequestDocsUserPend + (countDocs - countDocsPost - countDocsSent)
      End If






      If (flowCategory = "FABRICACAO") Then
         printDoc = ck_factory_display.Value
         totalFactoryFlowDocsPend = totalFactoryFlowDocsPend + docsPend
      End If

      If (flowCategory = "PRINCIPAL") Then
         printDoc = ck_show_prin.Value
         totalPrincipalFlowDocsPend = totalPrincipalFlowDocsPend + docsPend
      End If


      If (status = "FINALIZADO") Then
         printDocStatus = ck_show_completed.Value
      Else
         printDocStatus = True
      End If

      If (printDoc And printDocStatus) Then
         doc_list.AddItem XdbFactory.getData(respQuery, "id")
         doc_list.List(doc_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "title")
         doc_list.List(doc_list.ListCount - 1, 2) = flowCategory
         doc_list.List(doc_list.ListCount - 1, 3) = userName
         doc_list.List(doc_list.ListCount - 1, 4) = format(XdbFactory.getData(respQuery, "create_ate"), "dd/mm/yyyy")
         doc_list.List(doc_list.ListCount - 1, 5) = Xhelper.iff(status = "FINALIZADO", status, "PEND EXT. ( " & countDocPentExt & " )    REJ.  ( " & countDocsRejected & " )")
         doc_list.List(doc_list.ListCount - 1, 6) = Xhelper.iff(status = "FINALIZADO", status, countDocsInTheFlow)
         doc_list.List(doc_list.ListCount - 1, 7) = Xhelper.iff(status = "FINALIZADO", status, countDocsSent)
         doc_list.List(doc_list.ListCount - 1, 8) = Xhelper.iff(status = "FINALIZADO", status & " (" & countDocsPost & ") ", countDocsPost & " (" & countDocs & ")")
      End If
      respQuery.MoveNext
   Loop

   header_titles = Array("Requisição", "Solicitação", "Fluxo", "Requisitado por:", "Solicitado em:", "PEND. & REJ.", "NO FLUXO", "ENVIADOS", "CONCLUIDOS")
   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, doc_list_header)




   lb_pend.Caption = "F: " & totalFactoryFlowDocsPend & "  P: " & totalPrincipalFlowDocsPend & " (" & totalDocsPend & ")"
   lb_sent.Caption = totalRequestDocsSent
   lb_completed.Caption = totalRequestDocsPost
   lb_user_pend.Caption = totalRequestDocsUserPend
   lb_total.Caption = totalRequestDocs
   lb_post.Caption = totalRequestDocsSent + totalRequestDocsPost
   lb_rejected.Caption = totalRequestDocsRejected
   pendPercent = totalDocsPend / totalRequestDocs * 100
   lb_pend_percent.Caption = "(" & Round(pendPercent, 2) & " %)"
   lb_docs_schedules.Caption = totalScheduleDocs
   lb_total_request.Caption = countReqPend
   lb_req_completed.Caption = countReqCompleted
   lb_total_canceled.Caption = totalCanceledDocs
   lb_totalDocsInFlow.Caption = totalDocsInTheFlow
   lb_totalSentTo.Caption = totalDocsToSend
   lb_totalEng.Caption = totalDocLibEng
   lb_totalPendExtern.Caption = totalDocsPentExt

End Function


Private Sub btn_search_doc_Click()

   Dim flowType As String
   Dim SearchText As String
   Dim status As String

   status = select_status_filter.Value


   flowType = select_flow_type.List(select_flow_type.ListIndex, 0)
   SearchText = txt_search.Value

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_issue_request.searchRequestedDocument(SearchText, flowType, status)


   Call getRequestDocuments(respQuery)
End Sub

Private Function getDocumentsBySelectedResquest(ByVal requestId As String)
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_issue_request.get_request_docs(requestId)

   Call getRequestDocuments(respQuery)

End Function

Private Function getRequestDocuments(ByRef respQuery As ADODB.Recordset) As Boolean
   On Error GoTo ErrorHandler
      request_items_list.Clear
      grd_list.Clear

      Dim items() As Variant
      Dim rowIndex As Long
      Dim recordCount As Long
      Dim isSearchFoundRecords As Boolean

      recordCount = 0
      rowIndex = 1

      respQuery.MoveFirst

      Do Until respQuery.EOF
         recordCount = recordCount + 1
         respQuery.MoveNext
      Loop

      If recordCount > 0 Then
         respQuery.MoveFirst

         ReDim items(1 To recordCount, 1 To 14) ' Adjust the second dimension size based on the number of columns

         Call getRequestData(rowIndex, items, respQuery)

         Call fillRequestListWithDocs(items)

      Else
         Call Alert.Show("DOCUEMNTOS DA REQUISIÇÃO", "NENHUM RESULTADO ENCONTRADO", 2000)
      End If


    Exit Function

ErrorHandler:
      getRequestDocuments = False ' Return False If an error occurred during MoveFirst
      Call Alert.Show("PESQUISA CONCLUIDA", "NENHUM RESULTADO ENCONTRADO", 2000)
    Exit Function
End Function

'/*
'
'
'
'*/
Function RemoveEmptyItemsFromArray(arr() As Variant) As Variant()
   Dim result() As Variant
   Dim i As Long
   Dim j As Long
   Dim isEmptyRow As Boolean
   Dim numRows As Long

   numRows = UBound(arr, 1)
   ReDim result(1 To numRows, 1 To UBound(arr, 2))

   Dim outputIndex As Long
   outputIndex = 0

   For i = LBound(arr, 1) To UBound(arr, 1)
      isEmptyRow = True

      For j = LBound(arr, 2) To UBound(arr, 2)
         If Not IsEmpty(arr(i, j)) Then
            isEmptyRow = False
          Exit For
         End If
      Next j

      If Not isEmptyRow Then
         outputIndex = outputIndex + 1
         For j = LBound(arr, 2) To UBound(arr, 2)
            result(outputIndex, j) = arr(i, j)
         Next j
      End If
   Next i

   ReDim Preserve result(1 To outputIndex, 1 To UBound(arr, 2))

   RemoveEmptyItemsFromArray = result
End Function




Private Function getRequestData(ByRef rowIndex As Long, ByRef items() As Variant, ByRef respQuery As ADODB.Recordset)



   Dim grdSequence As String
   Dim grdCode As String
   Dim grdId As String
   Dim grdDate As String
   Dim requestStatus As String
   Dim grd As String
   Dim fullFilePath As String
   Dim responseUser As String
   Dim postStatus As String
   Dim obs As String
   Dim postUserMsg As String
   Dim docRevId As String
   Dim docId As String
   Dim revCode As String
   Dim docRequestId As String
   Dim reviewStatus As String
   Dim redNumber As String


   Dim respGrdQuery As ADODB.Recordset



   Do Until respQuery.EOF
      requestStatus = XdbFactory.getData(respQuery, "status")

      If (filterRequestedItem(requestStatus)) Then


         docRevId = XdbFactory.getData(respQuery, "rev_id")
         docId = XdbFactory.getData(respQuery, "id")
         revCode = UCase(Trim(XdbFactory.getData(respQuery, "rev_code_request")))
         docRequestId = XdbFactory.getData(respQuery, "doc_request_id")
         reviewStatus = XdbFactory.getData(respQuery, "rev_status")
         redNumber = XdbFactory.getData(respQuery, "red_number")

         If (docRevId = "") Then
            Call fixDocRequestId(docRequestId, docId, revCode)
         End If

         If (docRevId <> "") Then
            Set respGrdQuery = db_grd.getContractorGrdByDocReviewId(docRevId)
            grdCode = XdbFactory.getData(respGrdQuery, "grd_number")
            grdDate = XdbFactory.getData(respGrdQuery, "grd_date")
            grdId = XdbFactory.getData(respGrdQuery, "grd_id")

            grd = grdCode & " : " & format(grdDate, "dd-mm-yyyy")
         Else
            grd = ""
         End If

         fullFilePath = XdbFactory.getData(respQuery, "file_path")


         responseUser = XdbFactory.getData(respQuery, "response_user")
         postIn = XdbFactory.getData(respQuery, "post_in_date")
         postUserMsg = XdbFactory.getData(respQuery, "post_user_response_msg")


         postStatus = ""
         If (responseUser <> "") Then
            postStatus = Xhelper.iff(docRevId <> "" And requestStatus <> "LIB. ENG.", " [ LIB. ENG ] - ", "") & responseUser & Xhelper.iff(postIn <> "", " [ " & postIn & " ]", "")
         End If

         If (reviewStatus = "REJ" And requestStatus <> Constants.REJEITADO) Then
            requestStatus = Constants.SUBISTITUIR

         End If
         If (redNumber <> "") Then
            items(rowIndex, 1) = redNumber
            items(rowIndex, 2) = "Prioridade: [ " & XdbFactory.getData(respQuery, "priority") & " ]  - " & XdbFactory.getData(respQuery, "contract_item")
            items(rowIndex, REQ_TB_COLN_DOC_NUMBER + 1) = XdbFactory.getData(respQuery, "doc_number") & " REV: " & revCode & " TE: " & UCase(Trim(XdbFactory.getData(respQuery, "issue")))
            items(rowIndex, 4) = Left(XdbFactory.getData(respQuery, "description"), 71)
            items(rowIndex, 5) = requestStatus
            items(rowIndex, 6) = postStatus
            items(rowIndex, 7) = grd
            items(rowIndex, 8) = Xhelper.iff(postUserMsg <> "", "VER", "")
            items(rowIndex, 9) = ""
            items(rowIndex, 10) = fullFilePath
            items(rowIndex, REQ_TB_COLN_GRD_ID + 1) = grdId
            items(rowIndex, REQ_TB_COLN_DOC_ID + 1) = docId
            items(rowIndex, REQ_TB_COLN_DOC_REV_ID + 1) = docRevId
            items(rowIndex, 14) = docRequestId

            rowIndex = rowIndex + 1
         End If
      End If

      respQuery.MoveNext

   Loop


End Function

Private Function fillRequestListWithDocs(ByRef items() As Variant)



   Dim header_titles() As Variant
   header_titles = Array("RED", "Equipamento", "Documento", "Descrição", "Status", "Responsável - Postar Em", "GRD", "OBS")

   request_items_list.List = items

   Call Xform.SetColumnWidthsAndHeader(request_items_list, lblHidden, header_titles, request_items_header)

End Function


Private Function filterRequestedItem(ByVal status As String) As Boolean

   Dim response As Boolean


   Select Case status

    Case "CANCELADO"
      response = ck_show_item_canceled.Value

    Case "REJEITADO"

      response = ck_show_item_rejected.Value
    Case "CONCLUIDO"
      response = ck_show_item_post.Value

    Case Else
      response = True

   End Select


   filterRequestedItem = response

End Function


Private Function checkDocsInFactory(ByVal requestId As String)

   If (ck_check_factory) Then
      list_factory_docs_grd.Clear

      Dim respQuery As ADODB.Recordset
      Set respQuery = db_issue_request.getRequestDocsSentToFactory(requestId)



      Dim respGrdQuery As ADODB.Recordset


      Do Until respQuery.EOF

         list_factory_docs_grd.AddItem XdbFactory.getData(respQuery, "id")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 1) = XdbFactory.getData(respQuery, "contract_item")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 2) = XdbFactory.getData(respQuery, "description")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 3) = XdbFactory.getData(respQuery, "doc_number")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 4) = XdbFactory.getData(respQuery, "rev_code")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 5) = XdbFactory.getData(respQuery, "issue")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 6) = XdbFactory.getData(respQuery, "destiny")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 7) = XdbFactory.getData(respQuery, "grd_number")
         list_factory_docs_grd.List(list_factory_docs_grd.ListCount - 1, 8) = XdbFactory.getData(respQuery, "create_ate")
         respQuery.MoveNext
      Loop

      header_titles = Array("ID", "Equipamento", "Descrição", "Documento", "Rev.", "TE", "Destinatário", "GRD", "GRD DATA")
      Call Xform.SetColumnWidthsAndHeader(list_factory_docs_grd, lblHidden, header_titles, list_header_factor_docs_grd)
   End If
End Function

Private Sub btn_list_post_Click()

   Dim doc As Object

   Dim filePath As String
   Dim fileName As String

   Dim descktopFolder As String
   Dim fso As Object


   answer = MsgBox("Tem certeza que quer Emitir o(s) Documento(s) da Lista? " & vbCrLf & vbCrLf & "Total de Documentos a Serem Emitidos: (" & lb_total_list_docs.Caption & " )", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes) Then

      Set fso = CreateObject("Scripting.FileSystemObject")

      Dim objFolders As Object
      Set objFolders = CreateObject("WScript.Shell").SpecialFolders
      descktopFolder = objFolders("desktop") & "\"

      Dim desktopFolderName As String
      desktopFolderName = "emissao_" & format(Now(), "dd_mm_yyyy_hh_mm_ss_")


      Dim desktopFullPath As String
      desktopFullPath = descktopFolder & desktopFolderName & "\"

      If Not fso.FolderExists(desktopFullPath) Then

         fso.CreateFolder desktopFullPath

      End If

      Dim docRequestId  As String



      Dim logFileName As String
      logFileName = "emissao_" & Day(Now()) & "_" & Month(Now()) & "_" & Year(Now()) & ".txt"
      Dim logInfo As String

      For Each varKey In docToPostList.Keys()
         If (varKey <> "") Then
            Set doc = docToPostList(varKey)
            filePath = doc("fileFullPath")
            fileName = fso.GetFileName(filePath)
            fileNameSplited = Split(UCase(fileName), "_REV_")

            docRequestId = doc("docRequestId")
            Call file_helper.copyFilesWithCheckSum(filePath, desktopFullPath & fileName)
            Call act_doc_request.changeDocRequestStatus2(docRequestId, Constants.NO_FLUXO)

            logInfo = Day(Now()) & "_" & Month(Now()) & "_" & Year(Now()) & "  --  REQ_ID: " & doc("docRequestId") & " : " & doc("docStatus") & "  -  " & fileName & "   :   " & filePath
            Call h_text_file.Create(logInfo, logFileName)
         End If

         Next

         Call openDocReviewImportForm(desktopFullPath)
      End If

End Sub


Private Function openDocReviewImportForm(ByVal destiny As String)


   doc_post_request_form.Hide

   Load doc_review_import_form

   doc_review_import_form.import_files_folder_path = destiny
   doc_review_import_form.project_selected_id = project_selected_id
   doc_review_import_form.project_txt = project_txt
   doc_review_import_form.Show
   Set doc_review_import_form.docsRequested = CreateObject("Scripting.Dictionary")
   Set doc_review_import_form.docsRequested = docsSelectedToPost
   Call doc_review_import_form.readFilesAction


End Function

Private Sub btn_list_clear_Click()


   answer = MsgBox("Tem certeza que quer limpar a lista de documentos selecionados para Emitir?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then

      docToPostList.RemoveAll
      lb_total_list_docs.Caption = docToPostList.count
      Call Alert.Show("Lista a Emitir Excluída com Sucesso!!!", "", 1500)

   End If
End Sub

Private Sub btn_add_to_post_list_Click()
   addDocsToPostList
End Sub

Private Function addDocsToPostList()

   Dim postListIndex As String
   Dim coutntDocs As Long
   Dim docRequestIdIndex As String
   Dim docRequestIssueIndex As String
   Dim docIssueSplited() As String



   For i = 0 To request_items_list.ListCount - 1

      Dim doc As Object
      Set doc = CreateObject("Scripting.Dictionary")

      On Error Resume Next
      If request_items_list.Selected(i) = True Then

         requestStatus = request_items_list.List(i, 4)
         If (canChangeDocRequestStatus(requestStatus)) Then

            fileFullPath = request_items_list.List(i, 9)
            docRequestId = request_items_list.List(i, 13)
            doc("fileFullPath") = fileFullPath
            doc("docRequestId") = docRequestId
            doc("docStatus") = requestStatus


            docRequestIdIndex = "ID-" & docRequestId
            docId = request_items_list.List(i, 11)
            docRequestIssueIndex = "ISSUE-" & docId
            docIssueSplited = Split(request_items_list.List(i, 2), "TE:")
            docsSelectedToPost(docRequestIdIndex) = docId
            docsSelectedToPost(docRequestIssueIndex) = Trim(docIssueSplited(1))

            coutntDocs = coutntDocs + 1

            postListIndex = "n" & docId
            docToPostList.Add postListIndex, doc

         End If

      End If
   Next i
   lb_total_list_docs.Caption = docToPostList.count
   Call Alert.Show("Documentos Adicionados Com Sucesso: " & coutntDocs & " (" & docToPostList.count & ")", "", 1000)
End Function

Private Function getPostTempFolder() As String

   Dim tempFolder As String
   tempFolder = "cdoc_post_files_" & format(Now(), "yyyymmdd_hhmmss")

   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")

   Dim Temp_Fldr As String
   Temp_Fldr = fso.GetSpecialFolder(2)

   Dim cdoc_temp As String
   cdoc_temp = Temp_Fldr & "\" & tempFolder & "\"

   If Not fso.FolderExists(cdoc_temp) Then
      fso.CreateFolder cdoc_temp
   End If

   getPostTempFolder = cdoc_temp
End Function

Private Function canChangeDocRequestStatus(ByVal status As String) As Boolean
   Select Case status
    Case "CONCLUIDO", "HOLD", "CANCELADO", "ENVIADO", "REJEITADO"
      canChangeDocRequestStatus = False
      Call Alert.Show("Erro: Documento já em Fluxo", "", 2000)
    Case Else
      canChangeDocRequestStatus = True
   End Select
End Function



Private Sub request_items_list_Change()
   props_list.Clear
   grd_list.Clear


   Dim docId As String
   Dim docRevId As String

   For i = 0 To request_items_list.ListCount - 1
      On Error Resume Next
      If request_items_list.Selected(i) = True Then

         docId = request_items_list.List(i, 11)
         docRevId = request_items_list.List(i, 12)
         document_selected_id = docId
         Call getGrdsHandler(docRevId)
         Call getDocProps(docId)
         Call get_selected_doc_info(docId)
       Exit Sub
      End If
   Next i
End Sub

'/*
'
'
'
Private Sub btn_replaceDocumentFile_Click()

   Call Alert.Show("Iniciando a Substituição", "Aguarde", 1500)


   Dim docStatus As String
   Dim docFilePath As String
   Dim grd As String
   Dim docId As String
   Dim docRevId As String

   ' Iterate through the list items
   For i = 0 To request_items_list.ListCount - 1
      On Error Resume Next

      ' Check If the current item is selected
      If request_items_list.Selected(i) Then
         docStatus = request_items_list.List(i, REQ_TB_COLN_STATUS)

         ' Check If the document status requires replacement
         If docStatus = Constants.SUBISTITUIR Then
            ' Retrieve necessary information
            grd = request_items_list.List(i, REQ_TB_COLN_GRD_ID)
            docId = request_items_list.List(i, REQ_TB_COLN_DOC_ID)
            docRevId = request_items_list.List(i, REQ_TB_COLN_DOC_REV_ID)
            docFilePath = request_items_list.List(i, REQ_TB_COLN_DOC_FILE_PATH)

            ' Replace the document file
            act_replace_doc_not_reviewed.ReplaceDocFile docRevId, docFilePath


            If (grd <> "" And docRevId <> "") Then
               Call db_grd.softDeleteDocument(grd, docRevId)
            End If

         End If
      End If
   Next i

   ' Update documents from the selected request
   getDocsFromSelectedRequestHandler

   ' Display final alert
   Call Alert.Show("Substituição Finalizada", "", 2500)
End Sub


Private Function get_selected_doc_info(ByVal doc_id As String)
   If (doc_id <> "") Then
      fr_doc_titles.Enabled = True
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_documents.getDocumentById(doc_id)
      doc_name_txt.Value = XdbFactory.getData(respQuery, "name")
      doc_description_txt.Value = XdbFactory.getData(respQuery, "description")
      doc_number_acivate_txt.Value = XdbFactory.getData(respQuery, "doc_number")
      doc_pages_txt.Value = XdbFactory.getData(respQuery, "pages")
      Call set_doc_format_selected(XdbFactory.getData(respQuery, "doc_format"))
      Call set_doc_extension_selected(XdbFactory.getData(respQuery, "doc_extension"))
      Call setContractItemSelected(XdbFactory.getData(respQuery, "contract_item"))
      Call setCategorySelected(XdbFactory.getData(respQuery, "category"))
   Else
      clear_form_fields
   End If
End Function





Private Function setCategorySelected(ByVal category As String)

   For i = 0 To category_select.ListCount - 1
      If Trim(UCase(category_select.List(i, 1))) = Trim(UCase(category)) Then

         category_select.ListIndex = i
       Exit Function
      End If

   Next i
End Function
Private Function setContractItemSelected(contractItem As String)

   For i = 0 To contract_item_select.ListCount - 1
      If Trim(UCase(contract_item_select.List(i, 1))) = Trim(UCase(contractItem)) Then

         contract_item_select.ListIndex = i
       Exit Function
      End If

   Next i
End Function
Private Function set_doc_format_selected(format As String)

   For i = 0 To doc_format_select.ListCount - 1
      If doc_format_select.List(i, 1) = format Then

         doc_format_select.ListIndex = i
       Exit Function
      End If

   Next i
End Function

Private Function set_doc_extension_selected(extension As String)

   For i = 0 To extension_select.ListCount - 1
      If extension_select.List(i, 1) = extension Then

         extension_select.ListIndex = i
       Exit Function
      End If

   Next i

End Function
Private Function clear_form_fields()
   doc_name_txt.Value = ""
   doc_description_txt.Value = ""


   fr_doc_titles.Caption = "Documento Selecionado"
   fr_doc_titles.Enabled = False
End Function


'/*
'
'
'Get GRD's from the selected requested document
'
'*/
Private Function getGrdsHandler(ByVal revId As String)



   If (revId <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_grd.getGrdFromRequestDoc(revId)



      Do Until respQuery.EOF

         grd_list.AddItem XdbFactory.getData(respQuery, "id")
         grd_list.List(grd_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "grd_number") & "  -->  " & Left(XdbFactory.getData(respQuery, "destiny"), 21) & "  -->  " & format(XdbFactory.getData(respQuery, "create_ate"), "dd/mm/yyyy")

         respQuery.MoveNext
      Loop

   End If

End Function

Private Function getDocProps(ByVal docId As String)


   If (docId <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_document_props.getAll(docId)



      Do Until respQuery.EOF

         props_list.AddItem XdbFactory.getData(respQuery, "id")
         props_list.List(props_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name") & "  -->  " & XdbFactory.getData(respQuery, "value")


         respQuery.MoveNext
      Loop

   End If

End Function







Private Sub change_issue_btn_Click()

   Dim answer As Integer
   Dim requestStatus As String
   Dim docRequestId As String
   Dim doReviewId As String
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   data("issue") = issue_select.Value



   If (auth.is_authorized("SUPER_ADMIN")) Then

      answer = MsgBox("Tem certeza que quer Definir a Emissão?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

      If (answer = vbYes) Then

         For i = 0 To request_items_list.ListCount - 1
            On Error Resume Next
            If request_items_list.Selected(i) = True Then
               requestStatus = request_items_list.List(i, 4)
               If (canChangeDocRequestStatus(requestStatus)) Then
                  docRequestId = request_items_list.List(i, 13)
                  doReviewId = request_items_list.List(i, 12)
                  Call db_issue_request.updateRequestDocument(data, docRequestId)
                  If (doReviewId <> "") Then
                     Call db_documents.update_review_issue(doReviewId, issue_select.Value)
                  End If

               End If

            End If
         Next i
         getDocsFromSelectedRequestHandler
      End If
   Else
      MsgBox "Você não tem autorização para executar está operação", , "Autorização Negada"
   End If

End Sub



Private Sub btn_export_docs_Click()

   Dim answer As Integer

   answer = MsgBox("Quer exportar os dados da Requsição Selecionada?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
      If (requestIdSelected <> "") Then
         Dim fileName As String
         fileName = "requisicao_" & Day(Now()) & "_" & Month(Now()) & "_" & Year(Now()) & ".csv"
         Dim info As String
         info = ""
         Dim respQuery As ADODB.Recordset

         Set respQuery = db_issue_request.getRequestDocsWithProps(requestIdSelected)

         Do Until respQuery.EOF

            info = XdbFactory.getData(respQuery, "red") & ";"
            info = info & XdbFactory.getData(respQuery, "contract_item") & ";"
            info = info & XdbFactory.getData(respQuery, "doc_number") & ";"
            info = info & XdbFactory.getData(respQuery, "description") & ";"
            info = info & XdbFactory.getData(respQuery, "prop_value") & ";"
            info = info & XdbFactory.getData(respQuery, "rev_code") & ";"
            info = info & XdbFactory.getData(respQuery, "issue") & ";"
            info = info & XdbFactory.getData(respQuery, "category") & ";"
            info = info & XdbFactory.getData(respQuery, "doc_req_priority") & ";"
            info = info & XdbFactory.getData(respQuery, "prop_name")

            Call h_text_file.Create(info, fileName)

            info = ""
            respQuery.MoveNext
         Loop

         MsgBox "Exportação Finalizada: " & fileName
      End If
   End If
End Sub

Private Sub btn_update_doc_titles_Click()
   update_handler
   If (requestIdSelected <> "") Then
      Call getDocumentsBySelectedResquest(requestIdSelected)
      Call checkDocsInFactory(requestIdSelected)
   End If
End Sub


'/*
'
'
'Update document info handler
'
'*/
Private Function update_handler()


   Dim data As Object
   Dim answer As Integer
   Dim where As String


   Set data = CreateObject("Scripting.Dictionary")



   If (doc_name_txt.Value <> "" And doc_description_txt.Value <> "") Then

      answer = MsgBox("Quer Atualizar As Informações Do Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

      If (answer = vbYes) Then

         data("name") = doc_name_txt.Value
         data("description") = doc_description_txt.Value
         data("doc_format") = doc_format_select.List(doc_format_select.ListIndex, 1)
         data("doc_number") = doc_number_acivate_txt.Value
         data("project_contract_item_id") = contract_item_select.Value
         data("contract_item") = contract_item_select.List(contract_item_select.ListIndex, 1)
         data("category_id") = category_select.Value
         data("doc_extension") = extension_select.List(extension_select.ListIndex, 1)
         data("pages") = doc_pages_txt.Value


         where = "id='" & document_selected_id & "'"
         Call db_documents.update(data, where)

         Call Alert.Show("Dados Atualizados com Sucesso!!!", "", 2000)
         Call get_selected_doc_info(document_selected_id)
      End If

   Else

      MsgBox "Favor completar os daods", , "Dados incompletos"

   End If


End Function
