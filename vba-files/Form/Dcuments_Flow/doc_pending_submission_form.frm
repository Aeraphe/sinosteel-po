VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_pending_submission_form 
   Caption         =   "Situação Geral do Fluxo de  Documentos"
   ClientHeight    =   13635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20670
   OleObjectBlob   =   "doc_pending_submission_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_pending_submission_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Dcuments_Flow


Private project_selected_id
Private document_selected_id As String
Private document_review_selected_id As String
Private files_folder_path As String


Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)


End Sub

Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")


   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")
      read_files_btn.Enabled = True
      docs_import_fr.Enabled = True
      get_documents_handler
   Else
      read_files_btn.Enabled = False
      docs_import_fr.Enabled = False
   End If
End Sub



Private Function get_documents_handler()

   getDocumentsNotReturnedFromApproveFlow
   get_documents_replaced
   get_documents_not_sent_to_contractor
   check_if_has_last_review


End Function




Private Function getDocumentsNotReturnedFromApproveFlow()


   Frame3.Caption = "Documentos Agardando Retorno do Contratante"

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getDocumentsNotReturnFromApproveFlow(project_selected_id)
   doc_not_returned_list.Clear

   Dim countDocs As Long
   countDocs = 0

   Do Until respQuery.EOF




      doc_not_returned_list.AddItem XdbFactory.getData(respQuery, "doc_rev_id")

      doc_not_returned_list.List(doc_not_returned_list.ListCount - 1, 1) = "Verificar/Cobrar Retorno "
      doc_not_returned_list.List(doc_not_returned_list.ListCount - 1, 2) = UCase(XdbFactory.getData(respQuery, "name"))
      doc_not_returned_list.List(doc_not_returned_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "doc_number")
      doc_not_returned_list.List(doc_not_returned_list.ListCount - 1, 4) = "[ " & XdbFactory.getData(respQuery, "rev") & " ] : [ " & XdbFactory.getData(respQuery, "te") & " ]  : [ " & XdbFactory.getData(respQuery, "status") & " ]"
      doc_not_returned_list.List(doc_not_returned_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "grd_number")
      doc_not_returned_list.List(doc_not_returned_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "grd_date")
      doc_not_returned_list.List(doc_not_returned_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "grd_description")

      respQuery.MoveNext

      countDocs = countDocs + 1
   Loop

   Frame3.Caption = "Documentos Agardando Retorno do Contratante ( " & countDocs & " )"
   loadHeaderDocumentsNotReturnedFromApproveFlow

End Function



Private Function loadHeaderDocumentsNotReturnedFromApproveFlow()

   Dim header_titles As Variant

   header_titles = Array("ID", "Ação", "Enviado Para", "Nº Documentos", "REV x TE x STATUS", "GRD", "GRD Data", "GRD Nome")

   Call Xform.SetColumnWidthsAndHeader(doc_not_returned_list, lblHidden, header_titles, doc_not_returned_header)

End Function


Private Function get_documents_replaced()



   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.gel_all_doc_replaced_pend_to_submitting(project_selected_id)
   doc_replaced_list.Clear

   Do Until respQuery.EOF




      doc_replaced_list.AddItem XdbFactory.getData(respQuery, "doc_rev_id")

      doc_replaced_list.List(doc_replaced_list.ListCount - 1, 1) = "Enviar Mesma Revisão"
      doc_replaced_list.List(doc_replaced_list.ListCount - 1, 2) = UCase(XdbFactory.getData(respQuery, "name"))
      doc_replaced_list.List(doc_replaced_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "doc_number")
      doc_replaced_list.List(doc_replaced_list.ListCount - 1, 4) = "[ " & XdbFactory.getData(respQuery, "rev") & " ] : [ " & XdbFactory.getData(respQuery, "te") & " ] : [ " & XdbFactory.getData(respQuery, "status") & " ]"
      doc_replaced_list.List(doc_replaced_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "grd_number")
      doc_replaced_list.List(doc_replaced_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "grd_date")
      doc_replaced_list.List(doc_replaced_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "grd_description")

      respQuery.MoveNext


   Loop


   load_header_doc_pend_to_replaced

End Function


Private Function load_header_doc_pend_to_replaced()

   Dim header_titles As Variant

   header_titles = Array("ID", "Ação", "Enviado Para", "Nº Documentos", "REV x TE x STATUS", "GRD", "GRD Data", "GRD Nome")

   Call Xform.SetColumnWidthsAndHeader(doc_replaced_list, lblHidden, header_titles, doc_replaced_header_list)

End Function




Private Function get_documents_not_sent_to_contractor()

   Dim doc_review_id As String
   Dim doc_rev_check As String

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.get_documents_not_sent_to_contractor(project_selected_id)
   doc_to_post_on_ged_list.Clear

   Do Until respQuery.EOF



      doc_review_id = XdbFactory.getData(respQuery, "doc_rev_id")

      If (doc_review_id <> "") Then
         Dim respCheckQuery As ADODB.Recordset
         Set respCheckQuery = db_grd.get_doc_from_contractor_grd_recipient(project_selected_id, doc_review_id)

         doc_to_post_on_ged_list.AddItem doc_review_id
         doc_rev_check = XdbFactory.getData(respCheckQuery, "doc_rev_id")
         If (doc_review_id = doc_rev_check) Then

            doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 1) = "Confirmar Postado No GED"
         Else
            doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 1) = "Postar No GED"


         End If

         doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 2) = doc_review_id & " - " & UCase(XdbFactory.getData(respQuery, "name"))
         doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "doc_number")
         doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 4) = "[ " & XdbFactory.getData(respQuery, "rev") & " ] : [ " & XdbFactory.getData(respQuery, "te") & " ] : [ " & XdbFactory.getData(respQuery, "status") & " ]"
         doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "grd_number")
         doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "grd_date")
         doc_to_post_on_ged_list.List(doc_to_post_on_ged_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "grd_description")


      End If

      respQuery.MoveNext

   Loop


   load_header_doc_not_sent_to_contractor

End Function


Private Function load_header_doc_not_sent_to_contractor()

   Dim header_titles As Variant

   header_titles = Array("ID", "Ação", "Enviado Para", "Nº Documentos", "REV x TE x STATUS", "GRD", "GRD Data", "GRD Nome")

   Call Xform.SetColumnWidthsAndHeader(doc_to_post_on_ged_list, lblHidden, header_titles, doc_list_header)

End Function




Private Function check_if_has_last_review()

   Dim doc_review_id As String
   Dim doc_rev_check As String

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.check_if_has_last_doc_review(project_selected_id)
   doc_new_reviews_list.Clear

   Do Until respQuery.EOF




      doc_new_reviews_list.AddItem XdbFactory.getData(respQuery, "doc_rev_id")

      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 1) = "Enviar Nova Revisão"
      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 2) = UCase(XdbFactory.getData(respQuery, "name"))
      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "doc_number")
      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 4) = "[ " & XdbFactory.getData(respQuery, "rev") & " ] : [ " & XdbFactory.getData(respQuery, "te") & " ] : [ " & XdbFactory.getData(respQuery, "status") & " ]"
      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 5) = "ENVIAR ->> : [ " & XdbFactory.getData(respQuery, "last_review") & " ] : [ " & XdbFactory.getData(respQuery, "last_review_issue") & " ] : [ " & XdbFactory.getData(respQuery, "last_review_status") & " ]"
      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "grd_number")
      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "grd_date")
      doc_new_reviews_list.List(doc_new_reviews_list.ListCount - 1, 8) = XdbFactory.getData(respQuery, "grd_description")

      respQuery.MoveNext


   Loop

   load_header_next_review

End Function


Private Function load_header_next_review()

   Dim header_titles As Variant

   header_titles = Array("ID", "Ação", "Enviar Para", "Nº Documentos", "REV x TE x STATUS", "PROXIMA REV.", "GRD", "GRD Data", "GRD Nome")

   Call Xform.SetColumnWidthsAndHeader(doc_new_reviews_list, lblHidden, header_titles, doc_new_reviews_header_list)

End Function
