VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grd_create_form 
   Caption         =   "Gerador de GRDs"
   ClientHeight    =   12330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22650
   OleObjectBlob   =   "grd_create_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grd_create_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\GRD


Private selectedProjectId
Private isFormValid
Private search_doc_option As String

Private grd_doc_review_id_selected As String
Private doc_id_selected As String
Public projectIdSelected As String


Private Sub doc_list_Click()

End Sub

Private Sub grd_to_select_Change()
   selected_destiny = grd_to_select.ListIndex
End Sub



Private Sub load_docs_from_folder_btn_Click()

Dim sFolder As String

    sFolder = file_helper.open_folder_dialog

    
    If sFolder <> "" Then
      Call load_files_in_grd(sFolder)
    End If

End Sub


Private Function load_files_in_grd(folder_path As String)


   Set files_dict = CreateObject("Scripting.Dictionary")



   Dim doc_id As String


   grd_list.Clear
   Set files_dict = file_helper.get_files_from_folders(folder_path)

   Dim doc_code As String
   Dim doc_name As String
   Dim file_not_found As Integer
   Dim files_found_on_db As Integer
   files_found_on_db = 0
   file_not_found = 0

 

   For Each varKey In files_dict.Keys()
      If (varKey <> "count") Then


         file_name = files_dict(varKey)

         file = Split(UCase(file_name), "_REV_")
         On Error GoTo error_handler
         extension = Split(file(1), ".")

         next_rev = extension(0)


         Dim respQuery As ADODB.Recordset
         Set respQuery = db_documents.SearchLimit(selectedProjectId, Trim(UCase(file(0))), "doc_number")
         
         doc_code = XdbFactory.getData(respQuery, "doc_number")
         doc_id = XdbFactory.getData(respQuery, "id")
         rev_id = XdbFactory.getData(respQuery, "rev_id")

         If (doc_code <> "") Then

            doc_name = Left(XdbFactory.getData(respQuery, "name") & " - " & XdbFactory.getData(respQuery, "description"), 80)

            issue = respQuery.fields.item("issue")
            issue = Xhelper.iff(IsNull(issue), "-1", issue)

            last_rev = respQuery.fields.item("last_rev")
            last_rev = Xhelper.iff(IsNull(last_rev), "-1", last_rev)


      If (Len(doc_name) <= 130) Then
         description_size = Len(doc_name)
      Else
         description_size = title_size_txt.Value
      End If

            grd_list.AddItem rev_id
            grd_list.List(grd_list.ListCount - 1, 1) = UCase(doc_code)
            grd_list.List(grd_list.ListCount - 1, 2) = "[Rev: " & last_rev & "]   [TE: " & issue & "]"
            grd_list.List(grd_list.ListCount - 1, 3) = Left(doc_name, description_size)
            grd_list.List(grd_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "category")
            grd_list.List(grd_list.ListCount - 1, 5) = media_select.Value
            grd_list.List(grd_list.ListCount - 1, 6) = type_select.Value
            grd_list.List(grd_list.ListCount - 1, 7) = copies_txt.Value
            grd_list.List(grd_list.ListCount - 1, 8) = XdbFactory.getData(respQuery, "pages")
     

         
         Else
            Call Alert.Show("Documento não Cadastrado no Sistema", doc_code, 2000)
      
            file_not_found = file_not_found + 1

         End If


      End If
   Next '

    load_grd_docs_header


   Exit Function
error_handler:
   MsgBox "Erro: Documento Fora do Formato: " & varKey

End Function
Private Sub UserForm_Initialize()


   populate_destiny_select_handler
   getGRDMediaTypesHandler
   getGRDContentTypesHandler
   grd_date_txt.Value = Date
   isFormValid = False
   media_select.ListIndex = 0
   type_select.ListIndex = 0

   grd_description_txt.Value = auth.user_name & "_" & Now

End Sub


Private Sub UserForm_Activate()
   If (Not auth.is_logged) Then
      Unload Me
      Call Alert.Show("Favor Logar no Sistema", "", 2500)
   End If
End Sub



Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")
      destiny_fr.Enabled = True
      doc_search_fr.Enabled = True
      doc_list.Enabled = True

   End If
End Sub


Private Function getGRDMediaTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getMediaTypes()
   Call Shared_CommonSelectComp.Mount(media_select, respQuery, "name", "description")

End Function


Private Function populate_destiny_select_handler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getAll()
   Call Shared_CommonSelectComp.Mount(grd_to_select, respQuery, "id", "name")

End Function


Private Function getGRDContentTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getContentTypes()
   Call Shared_CommonSelectComp.Mount(type_select, respQuery, "name", "description")

End Function


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

   doc_list.Visible = False

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


   Dim header_titles As Variant
   header_titles = Array("id", "Documento", "Revisão", "Descrição", "Categoria", "Tipo", "Páginas")

   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, doc_list_header)
   doc_list.Visible = True

   Call Xhelper.waitMs(400)
   Unload UserFormAlert


End Function


Private Sub delete_doc_property_btn_Click()
   On Error Resume Next
   doc_properties_list.RemoveItem (doc_properties_list.ListIndex)
End Sub



Private Sub add_doc_btn_Click()

   Dim doc_rev_id As String

   Dim enable_fr As Boolean

   enable_fr = False

   For i = 0 To doc_list.ListCount - 1

      If doc_list.Selected(i) = True Then
         doc_rev_id = doc_list.List(i, 0)

         If (Not is_document_is_in_grd_list(doc_rev_id)) Then

            grd_list.AddItem doc_rev_id
            grd_list.List(grd_list.ListCount - 1, 1) = doc_list.List(i, 1) 'Doc number
            grd_list.List(grd_list.ListCount - 1, 2) = doc_list.List(i, 2) 'Doc Rev
            grd_list.List(grd_list.ListCount - 1, 3) = doc_list.List(i, 3) 'Description
            grd_list.List(grd_list.ListCount - 1, 4) = doc_list.List(i, 4) 'Category
            grd_list.List(grd_list.ListCount - 1, 5) = media_select.Value
            grd_list.List(grd_list.ListCount - 1, 6) = type_select.Value
            grd_list.List(grd_list.ListCount - 1, 7) = copies_txt.Value
            grd_list.List(grd_list.ListCount - 1, 8) = doc_list.List(i, 6) 'Pages

            enable_fr = True
         End If
      End If
   Next i

   load_grd_docs_header

   Call Alert.Show("Documentos Incluidos", "", 1500)

   grd_doc_list_fr.Enabled = enable_fr

End Sub


Public Function load_grd_docs_header()

   Dim header_titles As Variant
   header_titles = Array("id", "Documento", "Rev.", "Descrição", "Categoria", "Tipo", "Midia", "Copias", "Páginas")

   Call Xform.SetColumnWidthsAndHeader(grd_list, lblHidden, header_titles, grd_items_header_listbox)
End Function

'/*
'
'Check if the document is in grd list
'
'*/
Private Function is_document_is_in_grd_list(rev_id As String) As Boolean

   Dim i As Long

   For i = 0 To grd_list.ListCount - 1

      If (grd_list.List(i, 0) = rev_id) Then
         is_document_is_in_grd_list = True
         Exit Function
      End If

   Next i


   is_document_is_in_grd_list = False
End Function

Private Function isDocumentExistOnGRD(doc_review_id As String) As Boolean
   If (doc_review_id <> "") Then

      For i = 0 To grd_list.ListCount - 1

         If (doc_review_id = grd_list.List(i, 0)) Then
            isDocumentExistOnGRD = True
            Exit Function
         End If
      Next i
      isDocumentExistOnGRD = False
      Exit Function
   End If
   isDocumentExistOnGRD = True
End Function


Private Sub remove_btn_Click()
   For i = 0 To grd_list.ListCount - 1
      On Error Resume Next
      If grd_list.Selected(i) = True Then

         grd_list.RemoveItem (i)

      End If
   Next i


End Sub



Private Sub save_and_create_grd_btn_Click()

   Call GenerateGRDHandler(True)
   Unload Me
End Sub

Private Sub generete_grd_btn_Click()

   Me.Hide
   grd_simple_confirmatiom_form.Show

   If (grd_simple_confirmatiom_form.confirmation) Then
      Call GenerateGRDHandler(grd_simple_confirmatiom_form.options)
      Unload grd_simple_confirmatiom_form
   End If
   Unload Me
End Sub

Public Function GenerateGRDHandler(opt As Object)

   Dim grd_id  As String
    Dim data As Object
    Dim date_in_sqlite_format As String

   If (isGRDValid) Then
   
   If (grd_date_txt.Value <> "") Then
date_in_sqlite_format = DateHelpers.FormatDateToSQlite(grd_date_txt.Value)
End If

  
      Set data = CreateObject("Scripting.Dictionary")
      data("user_id") = auth.get_user_id
      data("recipent_id") = grd_to_select.Value
      data("issue_date") = date_in_sqlite_format
      data("description") = grd_description_txt.Value
      data("obs") = grd_obs.Value
      data("sequece_number") = getLastGRDNumberForSelectedRecipentHandler
      grd_id = db_grd.CreateGRD(data)
      Call InsertGRDItemsHandler(grd_id)
      Call action_grd.create_selected_grd_view(opt, grd_id)
  
   Else
      Call Alert.Show("Erro: Favor Preencher os Dados Faltantes", "", 2000)
   End If

End Function

Private Function getLastGRDNumberForSelectedRecipentHandler() As Long

   Dim grdNumber As Long
   grdNumber = 1
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getLastGRDFromRecipient(grd_to_select.Value)

   On Error Resume Next
   grdNumber = respQuery.fields("sequece_number").Value + 1

   getLastGRDNumberForSelectedRecipentHandler = grdNumber

End Function

Private Function InsertGRDItemsHandler(grd_id As String)

   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   For i = 0 To grd_list.ListCount - 1

If (grd_list.List(i, 0) <> "" And grd_id <> "") Then
      data("grd_id") = grd_id
      data("doc_rev_id") = grd_list.List(i, 0)
      data("doc_media_type") = grd_list.List(i, 5)
      data("doc_type") = grd_list.List(i, 6)
      data("doc_copies") = grd_list.List(i, 7)
      data("doc_pages") = grd_list.List(i, 8)

      Call db_grd.insertGRDDocuments(data)
End If

   Next i


End Function


Private Function isGRDValid() As Boolean

   If (grd_list.ListCount > 0 And grd_description_txt.Value <> "" And grd_to_select.Value <> "") Then
      isGRDValid = True
   Else
      isGRDValid = False
   End If

End Function


Private Sub grd_list_Change()
   populate_doc_review_list
End Sub




Private Function populate_doc_review_list()

   Dim docQuery As ADODB.Recordset
   Dim reviewQuery As ADODB.Recordset
  

   For i = 0 To grd_list.ListCount - 1

      If grd_list.Selected(i) = True Then

         grd_doc_review_id_selected = grd_list.List(i, 0)
         If (grd_doc_review_id_selected <> "") Then

            Set docQuery = db_documents.get_doc_by_review2(grd_doc_review_id_selected)
            Set reviewQuery = db_documents.getDocumentReviews(XdbFactory.getData(docQuery, "id"))

            Call Shared_CommonSelectComp.Mount(reviews_select, reviewQuery, "id", "rev_code")
         End If



         Exit Function
      End If
   Next i


End Function



Private Sub change_review_btn_Click()



   If (reviews_select.Value <> "") Then
      For i = 0 To grd_list.ListCount - 1

         If grd_list.Selected(i) = True Then

       

            grd_list.List(i, 0) = reviews_select.Value
            grd_list.List(i, 2) = reviews_select.List(reviews_select.ListIndex, 1)

            Call Alert.Show("Revisão Modificada com Sucesso!!!", "", 1700)
            Exit Sub
         End If
      Next i

   End If

End Sub

Private Sub update_grd_item_btn_Click()

   Dim is_doc_upadated As Boolean
   is_doc_upadated = False

   For i = 0 To grd_list.ListCount - 1

      If grd_list.Selected(i) = True Then



         grd_list.List(i, 5) = Xhelper.iff(media_select.Value <> "", media_select.Value, grd_list.List(i, 5))
         grd_list.List(i, 6) = Xhelper.iff(type_select.Value <> "", type_select.Value, grd_list.List(i, 6))
         grd_list.List(i, 7) = Xhelper.iff(copies_txt.Value <> "", copies_txt.Value, grd_list.List(i, 7))

         is_doc_upadated = True

      End If
   Next i

   If (is_doc_upadated) Then
      Call Alert.Show("Documentos Modificados", "", 1700)
   End If
End Sub





   Private Sub grd_list_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
      If (grd_doc_review_id_selected <> "") Then
         Call doc_selected_info_form.load_data(grd_doc_review_id_selected)
     
           doc_selected_info_form.Show
        End If
   End Sub
