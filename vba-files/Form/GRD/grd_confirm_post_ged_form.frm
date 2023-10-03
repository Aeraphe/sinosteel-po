VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grd_confirm_post_ged_form 
   Caption         =   "Confirmar documentos postado no GED (Vale)"
   ClientHeight    =   11730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23490
   OleObjectBlob   =   "grd_confirm_post_ged_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grd_confirm_post_ged_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\GRD

Private selectedProjectId As String
Private grd_id_selected As String
Private grd_number_selected As String
Private docRequestId As String

Private Const SENT = "ENVIADO PARA CONTRATANTE"
Private Const CONFIRM_RECEIVE = "CONFIRMADO O RECEBIMENTO PELO CONTRATANTE"




Private Sub Frame4_Click()

End Sub

Private Sub grd_itens_list_Click()

End Sub

Private Sub reject_document_btn_Click()
   Dim review_id As String


   For i = 0 To grd_itens_list.ListCount - 1

      If grd_itens_list.Selected(i) = True Then


         review_id = grd_itens_list.List(i, 0)
         If (reject_motive.Value <> "" And review_id <> "") Then
            Call action_reject_document.cdoc_contractor_reject(review_id, reject_motive.Value)

         End If
      End If

   Next i

End Sub

Private Sub UserForm_Initialize()
   ged_post_date_txt.Value = Date
   status_select.AddItem SENT
   status_select.AddItem CONFIRM_RECEIVE

End Sub


Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub

Private Sub search_btn_Click()


   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")
      populate_grd_list

   End If
End Sub


Private Function populate_grd_list()

   grd_list.Clear
   grd_itens_list.Clear

   Dim dataQuery As Object
   Set dataQuery = CreateObject("Scripting.Dictionary")

   dataQuery("PROJECT_ID") = selectedProjectId

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.get_grd_sent_to_project_contractor(dataQuery)


   Do Until respQuery.EOF

      grd_list.AddItem XdbFactory.getData(respQuery, "id")
      grd_list.List(grd_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "code") & XdbFactory.getData(respQuery, "sequece_number")
      grd_list.List(grd_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "description")
      grd_list.List(grd_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "issue_date")

      respQuery.MoveNext
   Loop

   Dim header_titles As Variant
   header_titles = Array("id", "GRD", "Descrição", "Data Emissão")

   Call Xform.SetColumnWidthsAndHeader(grd_list, lblHidden, header_titles, grd_header_listbox)

End Function

Private Sub grd_list_Change()
   grd_itens_list.Clear


   For i = 0 To grd_list.ListCount - 1

      If grd_list.Selected(i) = True Then
         grd_id_selected = grd_list.List(i, 0)

         If (grd_id_selected <> "") Then

            Call get_grd_docs_handler(grd_id_selected)
            grd_number_selected = grd_list.List(i, 1)
            grd_selected_number_lb.Caption = grd_number_selected
            fr_grd_doc_list.Enabled = True

            Exit Sub
         End If
      End If
   Next i

   fr_grd_doc_list.Enabled = False
End Sub


Private Function get_grd_docs_handler(grd_id As String)



   grd_itens_list.Clear

   Dim fr_enable As Boolean

   fr_enable = False
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")
   data("id") = grd_id

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getGRDItems(data)

   grd_itens_list.Clear

   Do Until respQuery.EOF

      docRequestId = XdbFactory.getData(respQuery, "rec_doc_id")

      grd_itens_list.AddItem XdbFactory.getData(respQuery, "doc_rev_id")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "doc_number") & "   [REV: " & XdbFactory.getData(respQuery, "rev_code") & "]    [Emissão: " & XdbFactory.getData(respQuery, "issue") & "]"
      grd_itens_list.List(grd_itens_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "name") & " - " & XdbFactory.getData(respQuery, "description")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "status")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "doc_type")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "doc_media_type")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "doc_copies")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "red")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 9) = docRequestId

      fr_enable = True
      respQuery.MoveNext
   Loop

   If (fr_enable) Then


      Dim header_titles As Variant
      header_titles = Array("id", "Documento", "Descrição", "Status", "Tipo", "Midia", "Copias", "RED")

      Call Xform.SetColumnWidthsAndHeader(grd_itens_list, lblHidden, header_titles, grd_items_header_listbox)


   End If
End Function

Private Sub btn_doc_sent_Click()

   Dim status As String
   Dim requestStatus As String

   If (status_select.Value <> "") Then

      If (status_select.Value = SENT) Then
         status = Constants.REVIEW_SATUS_EXP
         requestStatus = Constants.ENVIADO
      Else
         status = Constants.REVIEW_SATUS_POST
         requestStatus = Constants.CONCLUIDO
      End If


      Call changeStatusHandler(status, requestStatus)
   End If

End Sub





Private Function changeStatusHandler(ByVal status As String, ByVal rqStatus As String)

   Dim docsStatusUdated As Boolean
   docsStatusUdated = False

   If (grd_id_selected <> "") Then

      If (is_user_authorized(grd_id_selected)) Then
         docsStatusUdated = update_doc_review_status(status, rqStatus)

      Else


         MsgBox "Favor selecionar a GRD", , "Dados incompletos"

      End If

      If (docsStatusUdated) Then
         populate_grd_list
         MsgBox "Postagem dos Documentos Selecionados Confirmada com SUCESSO!!!", , "Confirmação de Postagem"
         grd_selected_number_lb.Caption = ""
      Else
         MsgBox "Selecione um Documento na Lista", , "Confirmação de Postagem"
      End If
   End If
End Function


Private Function update_doc_review_status(ByVal status As String, ByVal rqStatus As String) As Boolean

   is_doc_upadated = False

   If (ged_post_date_txt.Value <> "") Then
      Set data = CreateObject("Scripting.Dictionary")
      data("status") = status
      data("status_date") = DateHelpers.FormatDateToSQlite(Trim(ged_post_date_txt.Value))

      For i = 0 To grd_itens_list.ListCount - 1

         If grd_itens_list.Selected(i) = True Then

            sql_where = "id = '" & grd_itens_list.List(i, 0) & "'"
            Call db_documents.updateStatus(data, sql_where)
            docRequestId = grd_itens_list.List(i, 9)
            If (docRequestId <> "") Then
               Call act_doc_request.changeDocRequestStatus(docRequestId, rqStatus)
            End If
            is_doc_upadated = True

         End If

      Next i

      update_doc_review_status = is_doc_upadated
   Else
      MsgBox "Preencha a Data da postagem", , "Dados Faltantes"
   End If
End Function

Private Function is_user_authorized(grd_id As String) As Boolean

   Dim grdQuery As ADODB.Recordset
   Set grdQuery = db_grd.getById(grd_id)

   Dim user_role As String
   user_id = auth.get_user_id
   user_role = auth.get_user_role

   If (user_id = grdQuery.fields.item("user_id") Or user_role = "SUPER_ADMIN") Then

      is_user_authorized = True
      Exit Function
   Else
      Me.Hide
      Call Alert.Show("Voçê não tem permissão para efetuar esta operação", "", 3000)

      is_user_authorized = False
      Exit Function
   End If

End Function
