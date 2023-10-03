VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grd_admin_form 
   Caption         =   "Administrar GRDs"
   ClientHeight    =   13320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22740
   OleObjectBlob   =   "grd_admin_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grd_admin_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\GRD

Private selectedProjectId
Private search_type As String
Private recipient_id_selected As String
Private grd_id_selected As String
Private grd_doc_review_id_selected As String
Private grd_doc_id_selected As String
Private grd_number_selected As String
Private search_doc_option As String
Private doc_id_selected As String




Private Sub CommandButton1_Click()
   If (doc_id_selected <> "") Then
      Call doc_selected_info_form.load_data(doc_id_selected)
      Call Xhelper.waitMs(1000)
      doc_selected_info_form.Show
   End If
End Sub





Private Sub grd_itens_list_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   If (grd_doc_review_id_selected <> "") Then
      Call doc_selected_info_form.load_data(grd_doc_review_id_selected)

      doc_selected_info_form.Show
   End If
End Sub

Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub



Private Sub grd_to_select_Change()
   If (grd_to_select.Value <> "") Then
      recipient_id_selected = grd_to_select.List(grd_to_select.ListIndex, 0)
      selectedProjectId = grd_to_select.List(grd_to_select.ListIndex, 1)
   End If
End Sub

Private Sub UserForm_Initialize()


   Call shared_select_grd_to_comp.Mount(grd_to_select)
   populate_destiny_select_handler
   getGRDMediaTypesHandler
   getGRDContentTypesHandler

   media_add_select.ListIndex = 0
   type_add_select.ListIndex = 0

   Load doc_selected_info_form

End Sub




Private Function getGRDMediaTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getMediaTypes()
   Call Shared_CommonSelectComp.Mount(media_select, respQuery, "name", "description")
   respQuery.MoveFirst
   Call Shared_CommonSelectComp.Mount(media_add_select, respQuery, "name", "description")

End Function

Private Function getGRDContentTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getContentTypes()
   Call Shared_CommonSelectComp.Mount(type_select, respQuery, "name", "description")
   respQuery.MoveFirst
   Call Shared_CommonSelectComp.Mount(type_add_select, respQuery, "name", "description")

End Function



Private Function populate_destiny_select_handler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getAll()
   Call Shared_CommonSelectComp.Mount(grd_destiny_select, respQuery, "id", "name")

End Function

Private Sub search_doc_grd_Click()
   search_type = "SEARCH"
   search_grd_handler
End Sub

Private Sub search_all_grd_btn_Click()
   search_type = "SEARCH_ALL"
   search_grd_handler
End Sub


Private Function search_grd_handler()

   If (search_type <> "" And grd_to_select.Value <> "") Then
      Dim respQuery As ADODB.Recordset

      Select Case search_type
       Case "SEARCH_ALL"

         Set respQuery = db_grd.getAllGRDFromRecipient(grd_to_select.Value)
       Case "SEARCH"

         Set respQuery = db_grd.getAllGRDFromRecipient(grd_to_select.Value)
       Case Else

      End Select


      Call populate_grd_list_handler(respQuery)

   End If
End Function


Private Function populate_grd_list_handler(respQuery As ADODB.Recordset)

   grd_list.Clear
   Do Until respQuery.EOF

      grd_list.AddItem XdbFactory.getData(respQuery, "id")
      grd_list.List(grd_list.ListCount - 1, 1) = "GRD - " & XdbFactory.getData(respQuery, "sequece_number")
      grd_list.List(grd_list.ListCount - 1, 2) = CDate(XdbFactory.getData(respQuery, "issue_date"))
      grd_list.List(grd_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "confirmation_date")
      grd_list.List(grd_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "description")
      grd_list.List(grd_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "obs")

      respQuery.MoveNext
   Loop


   Dim header_titles As Variant

   header_titles = Array("id", "GRD", "Data", "Confirmação", "Descrição", "Observação")

   Call Xform.SetColumnWidthsAndHeader(grd_list, lblHidden, header_titles, grd_header_listbox)



End Function

Private Sub generate_btn_Click()


   Me.Hide
   grd_simple_confirmatiom_form.Show

   If (grd_simple_confirmatiom_form.confirmation) Then
      Call action_grd.create_selected_grd_view(grd_simple_confirmatiom_form.options, grd_id_selected)
      Unload grd_simple_confirmatiom_form
   End If
   Me.Show

End Sub


Private Sub grd_list_Change()

   reviews_select.Clear

   For i = 0 To grd_list.ListCount - 1

      If grd_list.Selected(i) = True Then
         grd_id_selected = grd_list.List(i, 0)
         If (grd_id_selected <> "") Then
            Call get_grd_docs_handler(grd_id_selected)
            grd_number_selected = grd_list.List(i, 1)
            grd_selected_number_lb.Caption = grd_number_selected
            gdr_update_fr.Enabled = True

            Exit Sub
         End If
      End If
   Next i
   grd_docs_list_fr.Enabled = False
   gdr_update_fr.Enabled = False
End Sub

Private Function get_grd_docs_handler(grd_id As String)

   Dim fr_enable As Boolean

   fr_enable = False
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")
   data("id") = grd_id

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getGRDItems(data)

   grd_itens_list.Clear

   Do Until respQuery.EOF

      doc_full_description = XdbFactory.getData(respQuery, "name") & " - " & XdbFactory.getData(respQuery, "description")

      grd_itens_list.AddItem XdbFactory.getData(respQuery, "doc_rev_id")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "doc_number") & "  [ " & XdbFactory.getData(respQuery, "status") & " ]"
      grd_itens_list.List(grd_itens_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "rev_code")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "issue")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 4) = doc_full_description
      grd_itens_list.List(grd_itens_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "doc_type")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "doc_media_type")
      grd_itens_list.List(grd_itens_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "doc_copies")

      fr_enable = True
      respQuery.MoveNext
   Loop

   If (fr_enable) Then


      Dim header_titles As Variant
      header_titles = Array("id", "Nº Documento", "Rev.", "TE", "Descrição", "Tipo", "Midia", "Copias")

      Call Xform.SetColumnWidthsAndHeader(grd_itens_list, lblHidden, header_titles, grd_items_header_listbox)

      grd_docs_list_fr.Enabled = fr_enable

   End If
End Function




Private Sub delete_grd_btn_Click()

   Dim grd As String
   grd = grd_list.List(grd_list.ListIndex, 1)
   Dim answer As Integer
   answer = MsgBox("Quer apagar a " & grd & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes And grd_id_selected <> "") Then

      Call delete_grd_handler
      Call search_grd_handler
   End If

End Sub

Private Function delete_grd_handler()


   If (is_user_authorized(grd_id_selected)) Then

      Call db_grd.delete(grd_id_selected)
      Call Alert.Show("GRD DELETADA COM SUCESSO!!!", "", 2000)

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



'/*
'
'Delete GRD Items btn action
'
'
'*/
Private Sub delete_grd_item_btn_Click()

   Dim grd As String
   grd = grd_list.List(grd_list.ListIndex, 1)
   Dim answer As Integer
   answer = MsgBox("Quer apagar os items selecionados da  [ " & grd & " ] ?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes And grd_id_selected <> "") Then

      Call delete_grd_items_handler

   End If

End Sub

Private Function delete_grd_items_handler()

   Dim deleted  As Boolean
   Dim docReview As ADODB.Recordset
   Dim docRequestId As String

   deleted = False


   If (is_user_authorized(grd_id_selected)) Then

      For i = 0 To grd_itens_list.ListCount - 1

         If grd_itens_list.Selected(i) = True Then
            grd_doc_review_id_selected = grd_itens_list.List(i, 0)
            If (grd_doc_review_id_selected <> "") Then


               Set docReview = db_documents.get_doc_by_review2(grd_doc_review_id_selected)
               docRequestId = XdbFactory.getData(docReview, "request_doc_id")

               Call db_grd.delete_document(grd_id_selected, grd_doc_review_id_selected)
               If (docRequestId <> "") Then
                  Call changeDocRequestStatus(docRequestId, "LIB. ENG")
               End If

               deleted = True


            End If
         End If
      Next i
      If (deleted) Then
         Call get_grd_docs_handler(grd_id_selected)
         Call Alert.Show("Docuemnto(s) apagados com Sucesso!!!", "", 2000)
      Else
         Call Alert.Show("Não foi possível apagar!!!", "", 2000)
      End If
   End If
End Function








Private Sub update_grd_btn_Click()
   Dim date_in_sqlite_format As String

   If (grd_id_selected <> "") Then


      If (is_user_authorized(grd_id_selected)) Then

         If (grd_date_txt.Value <> "") Then
            date_in_sqlite_format = DateHelpers.FormatDateToSQlite(grd_date_txt.Value)
         End If

         Call grd_update_handler("recipent_id", grd_destiny_select.Value, "Destinatário")
         Call grd_update_handler("obs", grd_obs.Value, "Observação")
         Call grd_update_handler("issue_date", date_in_sqlite_format, "Data de Envio")
         Call grd_update_handler("description", grd_description_txt.Value, "Descrição")
         Call grd_update_handler("sequece_number", grd_sequence_txt.Value, "Sequencial")

         Call search_grd_handler
         Call clear_grd_data_form
      End If

   Else
      MsgBox "Favor selecionar a GRD", , "Dados incompletos"
   End If


End Sub


Private Function clear_grd_data_form()
   grd_sequence_txt.Value = ""
   grd_description_txt.Value = ""
   grd_date_txt.Value = ""
   grd_destiny_select.ListIndex = -1
End Function

Private Function grd_update_handler(prop As String, prop_value As Variant, change_type As String)


   If (Not IsNull(prop_value)) Then
      If (prop <> "" And prop_value <> "") Then


         UserFormAlert.Label1.Caption = "Atualizando GRD"
         UserFormAlert.Show
         UserFormAlert.Repaint

         Dim data As Object
         Set data = CreateObject("Scripting.Dictionary")
         data(prop) = prop_value


         Dim where As String

         where = "id='" & grd_id_selected & "'"
         Call db_grd.update(data, where)
         changed = True


         Unload UserFormAlert

         If (changed) Then
            Call Alert.Show("GRD Modificados com Sucesso!!!", "[ " & change_type & " ]", 2500)

         End If
      End If
   End If
End Function



Private Sub update_doc_btn_Click()

   If (grd_id_selected <> "") Then


      If (is_user_authorized(grd_id_selected)) Then
         Dim changed  As Boolean

         changed = False


         For i = 0 To grd_itens_list.ListCount - 1

            If grd_itens_list.Selected(i) = True Then
               grd_doc_review_id_selected = grd_itens_list.List(i, 0)
               If (grd_doc_review_id_selected <> "") Then


                  Call update_grd_doc_selected("doc_media_type", media_select.Value, "Tipo de Midia")
                  Call update_grd_doc_selected("doc_type", type_select.Value, "Tipo de Documento")
                  Call update_grd_doc_selected("doc_copies", copies_txt.Value, "Número de Copias")


                  changed = True


               End If
            End If
         Next i


         If (Not changed) Then
            Call Alert.Show("Erro: Selecione um item a ser modificado!!!", "", 2000)
         Else
            Call get_grd_docs_handler(grd_id_selected)
         End If

      End If

   Else
      MsgBox "Favor selecionar a GRD", , "Dados incompletos"
   End If
End Sub



Private Function update_grd_doc_selected(prop As String, prop_value As Variant, change_type As String)


   If (Not IsNull(prop_value)) Then
      If (prop <> "" And prop_value <> "") Then


         UserFormAlert.Label1.Caption = "Atualizando Doumento da GRD"
         UserFormAlert.Show
         UserFormAlert.Repaint

         Dim data As Object
         Set data = CreateObject("Scripting.Dictionary")
         data(prop) = prop_value


         Dim where As String

         where = "grd_id='" & grd_id_selected & "' AND  doc_rev_id='" & grd_doc_review_id_selected & "'"
         Call db_grd.update_document(data, where)
         changed = True


         Unload UserFormAlert

         If (changed) Then
            Call Alert.Show("Documento(s) Modificados com Sucesso!!!", "[ " & change_type & " ]", 2500)

         End If
      End If
   End If
End Function



Private Sub grd_itens_list_Change()


   Dim docQuery As ADODB.Recordset
   Dim respQuery As ADODB.Recordset

   If (grd_id_selected <> "") Then
      For i = 0 To grd_itens_list.ListCount - 1

         If grd_itens_list.Selected(i) = True Then

            grd_doc_review_id_selected = grd_itens_list.List(i, 0)
            If (grd_doc_review_id_selected <> "") Then

               Set docQuery = db_documents.get_doc_by_review2(grd_doc_review_id_selected)
               doc_id_selected = docQuery.fields.item("id")
               Set respQuery = db_documents.getDocumentReviews(doc_id_selected)
               Call Shared_CommonSelectComp.Mount(reviews_select, respQuery, "id", "rev_code")

               Exit Sub
            End If
         End If
      Next i
   End If

End Sub


Private Sub change_review_btn_Click()



   If (reviews_select.Value <> "" And Not is_document_is_in_grd_list(reviews_select.Value)) Then
      For i = 0 To grd_itens_list.ListCount - 1

         If grd_itens_list.Selected(i) = True Then



            Call update_grd_doc_selected("doc_rev_id", reviews_select.Value, "Modificado a Revisão")

            Call get_grd_docs_handler(grd_id_selected)
            Call Alert.Show("Revisão Modificada com Sucesso!!!", "", 1700)

            Exit Sub
         End If
      Next i

   End If

End Sub








Private Sub search_sec_title_btn_Click()
   search_doc_option = "description"
   Call SearchDocumentHandler(search_doc_option)
End Sub


Private Sub search_doc_btn_Click()
   search_doc_option = "name"
   Call SearchDocumentHandler(search_doc_option)
End Sub


Private Sub search_doc_sinosteel_btn_Click()
   search_doc_option = "sinosteel_doc_number"
   Call SearchDocumentHandler(search_doc_option)
End Sub


Private Sub search_doc_supplier_btn_Click()

   search_doc_option = "doc_number"
   Call SearchDocumentHandler(search_doc_option)

End Sub



Private Function SearchDocumentHandler(tb As String)

   Load UserFormAlert
   UserFormAlert.Label1.Caption = "Buscando Documentos"
   UserFormAlert.Show

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.SearchLastDocumentReview(selectedProjectId, search_doc_txt.Value, tb)

   Dim doc_number As String

   doc_list.Clear

   Do Until respQuery.EOF

      doc_number = XdbFactory.getData(respQuery, "doc_number")

      doc_list.AddItem XdbFactory.getData(respQuery, "rev_id")
      description = XdbFactory.getData(respQuery, "name") & " -- " & XdbFactory.getData(respQuery, "description")

      If (Len(description) <= 130) Then
         description_size = Len(description)
      Else
         description_size = title_size_txt.Value
      End If

      doc_list.List(doc_list.ListCount - 1, 1) = doc_number
      doc_list.List(doc_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "last_rev")
      doc_list.List(doc_list.ListCount - 1, 3) = Left(description, description_size)
      doc_list.List(doc_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "category")
      doc_list.List(doc_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "doc_type")
      doc_list.List(doc_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "pages")

      UserFormAlert.labelInfo.Caption = doc_number
      UserFormAlert.Repaint

      respQuery.MoveNext
   Loop

   UserFormAlert.Label1.Caption = "Busca Concluida"
   UserFormAlert.labelInfo.Caption = ""
   UserFormAlert.Repaint

   Call Xhelper.waitMs(2500)



   Dim header_titles As Variant
   header_titles = Array("id", "Documento", "Revisão", "Descrição", "Categoria", "Tipo", "Páginas")

   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, docs_header_listbox)

   Call Xhelper.waitMs(400)
   Unload UserFormAlert


End Function



Private Sub add_doc_btn_Click()
   Dim answer As Integer


   answer = MsgBox("Deseja incluir os items na GRD selecionada", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
      If (grd_id_selected <> "") Then


         If (is_user_authorized(grd_id_selected)) Then
            Call InsertGRDItemsHandler(grd_id_selected)
            Call Alert.Show("Documento(s) inserido na " & grd_number_selected, "", 2000)
            Call get_grd_docs_handler(grd_id_selected)
         End If

      End If
   End If

End Sub


'/*
'
'Check if the document is in grd list
'
'*/
Private Function is_document_is_in_grd_list(rev_id As String) As Boolean

   Dim i As Long

   For i = 0 To grd_itens_list.ListCount - 1
      grd_item_rev_id = grd_itens_list.List(i, 0)
      If (grd_item_rev_id = rev_id) Then
         is_document_is_in_grd_list = True
         Exit Function
      End If

   Next i


   is_document_is_in_grd_list = False
End Function


Private Function InsertGRDItemsHandler(grd_id As String)


   Dim doc_rev_id As String
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")



   For i = 0 To doc_list.ListCount - 1

      If doc_list.Selected(i) = True Then

         doc_rev_id = doc_list.List(i, 0)
         If (Not is_document_is_in_grd_list(doc_rev_id)) Then

            data("grd_id") = grd_id
            data("doc_rev_id") = doc_rev_id
            data("doc_media_type") = media_add_select.Value
            data("doc_type") = type_add_select.Value
            data("doc_copies") = copies_add_txt.Value

            Call db_grd.insertGRDDocuments(data)
         Else

            Call Alert.Show("Documento já exite na GRD", "", 3000)

         End If
      End If

   Next i


End Function
