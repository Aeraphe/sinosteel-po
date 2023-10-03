VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_mult_update_props_form 
   Caption         =   "Adicionar/Modificar Categoria, Disciplina ou Item do Contrato"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23430
   OleObjectBlob   =   "doc_mult_update_props_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_mult_update_props_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents

Public project_selected_id
Private search_doc_type As String









Private Sub UserForm_Initialize()
   Call Shared_DocCategorySelectComp.Mount(doc_category_select)
   getDisciplineHandler
   getDocExtensionstHandler
   getDocCodeTypestHandler
   getDocFormatstHandler
   populate_supplier_select_handler

End Sub
Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub





Private Sub change_btn_Click()


   Dim doc_format As String
   On Error Resume Next
   doc_format = UCase(doc_format_select.List(doc_format_select.ListIndex, 1))

   Dim doc_extension As String
   doc_extension = ""
   On Error Resume Next
   doc_extension = UCase(extension_select.List(extension_select.ListIndex, 1))
   
    Dim supplier_id As String
   supplier_id = ""
   On Error Resume Next
   supplier_id = UCase(supplier_select.List(supplier_select.ListIndex, 0))

   Dim contract_item As String
   contract_item = Trim(UCase(contract_item_txt.Value))


   Dim answer As Integer

   answer = MsgBox("Quer Atualizar os Documentos?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes) Then

      Call upadte_doc_propertie("doc_format", doc_format, "Formato")
      Call upadte_doc_propertie("category_id", doc_category_select.Value, "Categoria")
      Call upadte_doc_propertie("discipline_id", discipline_select.Value, "Disciplina")
      Call upadte_doc_propertie("doc_type_code", doc_code_seletc.Value, "Código")
      Call upadte_doc_propertie("doc_extension", doc_extension, "Extenção")
      Call upadte_doc_propertie("supplier_id", supplier_id, "Fornecedor")
      Call upadte_doc_propertie("contract_item", UCase(select_contract_item.List(select_contract_item.ListIndex, 1)), "Item do Contrato")
      Call upadte_doc_propertie("project_contract_item_id", select_contract_item.Value, "Item do Contrato")
    
      Call SearchDocumentHandler(search_doc_type)

   End If

End Sub


Private Function upadte_doc_propertie(prop As String, prop_value As String, change_type As String)


If (auth.is_authorized("SUPER_ADMIN")) Then
   If (prop <> "" And prop_value <> "") Then

      Dim selectedDocumentId As String
      Dim document_number As String
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      UserFormAlert.Label1.Caption = "Atualizando Documentos"
      UserFormAlert.Show
      UserFormAlert.Repaint


      For i = 0 To doc_list.ListCount - 1
         If doc_list.Selected(i) = True Then

            selectedDocumentId = doc_list.List(i, 0)
            document_number = doc_list.List(i, 1)
            data(prop) = prop_value
            Call update_handler(data, selectedDocumentId, document_number)
            changed = True

         End If
      Next i

      Unload UserFormAlert

      If (changed) Then
         Call Alert.Show("Documentos Modificados com Sucesso!!!", "[ " & change_type & " ]", 2500)

      End If

   End If
   
   Else
   Call Alert.Show("Voçê não tem permissão para efetuar está operação", "", 2000)
   
End If

End Function


Private Function update_handler(data As Object, doc_id As String, document_number As String)

   Dim where As String

   If (doc_id <> "") Then

      where = "id='" & doc_id & "'"
      Call db_documents.update(data, where)
      UserFormAlert.labelInfo.Caption = "Doc: " & document_number
      UserFormAlert.Repaint
   Else
      MsgBox "Favor completar os daods", , "Dados incompletos"
   End If

End Function

Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")
      search_doc_fr.Enabled = True
      contract_item_fr.Enabled = True
      sub_folder_fr.Enabled = True
      change_btn.Enabled = True
      btn_move_project_files.Enabled = True
      btn_move_comments_project_files.Enabled = True
      getContractItemsHandler
      Else
      
      search_doc_fr.Enabled = False
      change_btn.Enabled = False
       contract_item_fr.Enabled = False
        sub_folder_fr.Enabled = False
        btn_move_project_files.Enabled = False
         btn_move_comments_project_files.Enabled = False

   End If
End Sub


Private Function getContractItemsHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_projects.get_contract_items(project_selected_id)

   select_contract_item.Clear
   Call Shared_CommonSelectComp.Mount(select_contract_item, respQuery)
End Function

Private Function getDocFormatstHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllDocFormats()

   doc_format_select.Clear
   Call Shared_CommonSelectComp.Mount(doc_format_select, respQuery)
End Function

Private Function getDocCodeTypestHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllDocCodes()

   doc_code_seletc.Clear
   Call Shared_CommonSelectComp.Mount(doc_code_seletc, respQuery)
End Function




Private Function getDocExtensionstHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllExtensions()

   extension_select.Clear
   Call Shared_CommonSelectComp.Mount(extension_select, respQuery)
End Function



Private Function getDisciplineHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = database_discipline.getAll()

   discipline_select.Clear
   Call Shared_CommonSelectComp.Mount(discipline_select, respQuery)
End Function


Private Function populate_supplier_select_handler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = SuppliersDataBase.getAll()

   supplier_select.Clear
   Call Shared_CommonSelectComp.Mount(supplier_select, respQuery)
End Function


Private Sub search_doc_description_btn_Click()
   search_doc_type = "description"
   Call SearchDocumentHandler(search_doc_type)
End Sub

Private Sub search_doc_btn_Click()
   search_doc_type = "name"
   Call SearchDocumentHandler(search_doc_type)
End Sub


Private Sub search_doc_sinosteel_btn_Click()

   search_doc_type = "sinosteel_doc_number"
   Call SearchDocumentHandler(search_doc_type)
End Sub


Private Sub search_doc_supplier_btn_Click()
   search_doc_type = "doc_number"
   Call SearchDocumentHandler(search_doc_type)
End Sub



Private Function SearchDocumentHandler(tb As String)

   doc_list.Clear
   UserFormAlert.Show
   UserFormAlert.Label1.Caption = "Buscando Documentos"

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.search(project_selected_id, search_doc_txt.Value, tb)

   Dim doc_count As Long
   doc_count = 0
   total_docs_lb.Caption = 0

   doc_list.Visible = False
   Do Until respQuery.EOF

  

      description = XdbFactory.getData(respQuery, "name") & " -- " & XdbFactory.getData(respQuery, "description")

      If (Len(description) <= 80) Then
         description_size = Len(description)
      Else
         description_size = title_size_txt.Value
      End If
      doc_number = XdbFactory.getData(respQuery, "doc_number")
      UserFormAlert.labelInfo.Caption = doc_number
      UserFormAlert.Repaint

      new_description = Left(description, description_size)
    
      doc_list.AddItem XdbFactory.getData(respQuery, "id")
      doc_list.List(doc_list.ListCount - 1, 1) = doc_number
      doc_list.List(doc_list.ListCount - 1, 2) = new_description
      doc_list.List(doc_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "doc_type_code") & " - " & XdbFactory.getData(respQuery, "category")
      doc_list.List(doc_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "discipline")
      doc_list.List(doc_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "folder")
      doc_list.List(doc_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "contract_item")
      doc_list.List(doc_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "doc_extension") & " - " & XdbFactory.getData(respQuery, "doc_format") & " -  FOR: " & XdbFactory.getData(respQuery, "supplier")
     

      doc_count = doc_count + 1
      respQuery.MoveNext
   Loop
   total_docs_lb.Caption = doc_count
  
     UserFormAlert.Label1.Caption = "Busca Finalizada"
       UserFormAlert.labelInfo.Caption = "Documento(s) Encontrado(s): " & doc_count
     UserFormAlert.Repaint

   Dim header_titles As Variant
   header_titles = Array("id", "Documento", "Descrição", "Código - Categoria", "Disciplina", "Pasta (Opcional)", "Item Contrato", "Info")

   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, doc_list_header)
   doc_list.Visible = True
 
   Xhelper.waitMs (2000)
 
   Unload UserFormAlert


End Function

Private Sub btn_move_project_files_Click()
    Dim folder_path As String
    Dim response As VbMsgBoxResult
    
    folder_path = file_helper.open_folder_dialog

    ' Ask for confirmation
    response = MsgBox("Quer organizar as pastas dos documentos Emitidos?", vbYesNo + vbQuestion, "Confirmar")
    
    If response = vbYes Then
        If (folder_path <> "" And project_selected_id <> "") Then
            Call action_organize_project_files.start(project_selected_id, folder_path, "SENT")
        End If
    End If
End Sub

Private Sub btn_move_comments_project_files_Click()
    Dim folder_path As String
    Dim response As VbMsgBoxResult

    folder_path = file_helper.open_folder_dialog

    ' Ask for confirmation
    response = MsgBox("Quer organizar as pastas dos documentos Comentados?", vbYesNo + vbQuestion, "Confirmar")

    If response = vbYes Then
        If (folder_path <> "" And project_selected_id <> "") Then
            Call action_organize_project_files.start(project_selected_id, folder_path, "COMMENTS_WITH_DOC_ID")
        End If
    End If
End Sub

