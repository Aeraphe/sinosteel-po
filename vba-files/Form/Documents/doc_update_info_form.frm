VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_update_info_form 
   Caption         =   "Modificar informaçõesdo Documento"
   ClientHeight    =   12585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19770
   OleObjectBlob   =   "doc_update_info_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_update_info_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents


Private selectedProjectId
Private selectedDocumentId
Private search_doc_type As String
Private selected_doc_cat_id As String




Private Sub change_cat_btn_Click()

   Dim answer As Integer
   answer = MsgBox("Quer Modificar a Categoria do Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes And doc_category_select.Value <> "") Then
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      data("category_id") = doc_category_select.Value


      Dim where As String
      where = "id='" & selectedDocumentId & "'"
      Call db_documents.update(data, where)
      Call Alert.Show("Categoria Modificada com Sucesso", "", 2500)
      get_doc_categorie_handler

   End If
End Sub

Private Sub change_discipline_btn_Click()
   Dim answer As Integer
   answer = MsgBox("Quer Modificar a Disciplina do Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes And discipline_select.Value <> "") Then
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      data("discipline_id") = discipline_select.Value


      Dim where As String
      where = "id='" & selectedDocumentId & "'"
      Call db_documents.update(data, where)
      Call Alert.Show("Disciplina Modificada com Sucesso", "", 2500)
      get_doc_discipline_handler

   End If
End Sub

Private Sub change_doc_code_btn_Click()
   Load util_doc_code_select_form
   util_doc_code_select_form.code_label = doc_code_txt.Value
   util_doc_code_select_form.Show
   If (util_doc_code_select_form.doc_code_seletc.Value <> "") Then
      doc_code_txt.Value = UCase(util_doc_code_select_form.doc_code_seletc.List(util_doc_code_select_form.doc_code_seletc.ListIndex, 1))
   End If
End Sub

Private Sub change_doc_extension_btn_Click()
   Load util_extension_select_form
   util_extension_select_form.extension_label = doc_extension_txt.Value
   util_extension_select_form.Show
   If (util_extension_select_form.extension_select.Value <> "") Then
      doc_extension_txt.Value = UCase(util_extension_select_form.extension_select.List(util_extension_select_form.extension_select.ListIndex, 1))
   End If

End Sub

Private Sub change_doc_formt_btn_Click()
   Load util_doc_format_select_form
   util_doc_format_select_form.format_label = doc_formt_txt.Value
   util_doc_format_select_form.Show
   If (util_doc_format_select_form.doc_format_select.Value <> "") Then
      doc_formt_txt.Value = UCase(util_doc_format_select_form.doc_format_select.List(util_doc_format_select_form.doc_format_select.ListIndex, 1))
   End If
End Sub



Private Sub delete_doc_property_btn_Click()
   Dim answer As Integer
   answer = MsgBox("Quer Apagar a Propriedade do Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes And doc_properties_list.Value <> "") Then
      Call db_document_props.delete(selectedDocumentId, doc_properties_list.Value)
      Call Alert.Show("Propriedade Excluida com Sucesso", "", 2500)
      load_doc_props_handler
   End If
End Sub


Private Sub add_doc_property_btn_Click()
   Dim answer As Integer
   answer = MsgBox("Quer Adicionar a Propriedade ao Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes And doc_property_select.Value <> "" And doc_property_value.Value <> "") Then
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")
      data("document_id") = selectedDocumentId
      data("property_id") = doc_property_select.Value
      data("value") = doc_property_value.Value
      Call db_document_props.Create(data)
      Call Alert.Show("Propriedade Adicionada com Sucesso", "", 2500)
      load_doc_props_handler
   End If
End Sub


Private Sub delete_equipament_btn_Click()
   Dim answer As Integer
   answer = MsgBox("Quer Apagar o Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes And equipament_list.Value <> "") Then
      Call db_documents_equipaments.delete(selectedDocumentId, equipament_list.Value)
      Call Alert.Show("Equipamento Excluido com Sucesso", "", 2500)
      load_doc_equipament_handler
   End If
End Sub

Private Sub add_equipament_btn_Click()

   Dim answer As Integer
   answer = MsgBox("Quer Adicionar o Equipamento ao Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes And equipament_select.Value <> "") Then
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")
      data("document_id") = selectedDocumentId
      data("equipament_id") = equipament_select.Value
      Call db_documents_equipaments.Create(data)
      Call Alert.Show("Equipamento Adicionado com Sucesso", "", 2500)
      load_doc_equipament_handler
   End If
End Sub

Private Sub search_project_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")
      search_doc_frame.Enabled = True
      doc_inf_frame.Enabled = True

   Else
      search_doc_frame.Enabled = False
      doc_inf_frame.Enabled = False
   End If
End Sub

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


   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.search(selectedProjectId, doc_txt.Value, tb)

   doc_list.Clear
   Do Until respQuery.EOF

      doc_list.AddItem XdbFactory.getData(respQuery, "id")

      doc_list.List(doc_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "doc_number")
      doc_list.List(doc_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "name")
      doc_list.List(doc_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "category")
      doc_list.List(doc_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "doc_type")


      respQuery.MoveNext
   Loop

   Call SetColumnWidths(doc_list, lblHidden)
End Function



Private Sub doc_list_Click()
   selectedDocumentId = doc_list.Value
   select_doc_handler
   load_doc_equipament_handler
   load_doc_props_handler
   get_doc_categorie_handler
   get_doc_discipline_handler

End Sub


Private Function select_doc_handler()

   If (doc_list.Value <> "") Then

      Dim respQuery As ADODB.Recordset
      Set respQuery = db_documents.getDocumentById(doc_list.Value)


      sinosteel_number_txt.Value = XdbFactory.getData(respQuery, "sinosteel_doc_number")
      doc_number_txt.Value = XdbFactory.getData(respQuery, "doc_number")
      doc_name_txt.Value = XdbFactory.getData(respQuery, "name")
      doc_description_txt.Value = XdbFactory.getData(respQuery, "description")
      doc_extension_txt.Value = XdbFactory.getData(respQuery, "doc_extension")
      doc_code_txt.Value = XdbFactory.getData(respQuery, "doc_type_code")
      doc_formt_txt.Value = XdbFactory.getData(respQuery, "doc_format")
      doc_pages_txt.Value = XdbFactory.getData(respQuery, "pages")
      obs_txt.Value = XdbFactory.getData(respQuery, "obs")

      doc_inf_frame.Enabled = True
      adm_equipament_frame.Enabled = True
      adm_doc_prop_frame.Enabled = True
      adm_category_frame.Enabled = True
      adm_discipline_frame.Enabled = True

   Else

      doc_inf_frame.Enabled = False
      adm_equipament_frame.Enabled = False
      adm_doc_prop_frame.Enabled = False
      adm_category_frame.Enabled = False
      adm_discipline_frame.Enabled = False
   End If
End Function



'/*
'[EVENT]
'
'
'Btn to bootstrap upadate the document info
'
'*/
Private Sub update_btn_Click()
   update_handler
   select_doc_handler
End Sub


'/*
'
'
'Update document info handler
'
'*/
Private Function update_handler()


   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Dim where As String



   If (doc_list.Value <> "" And validate_form) Then
      Dim answer As Integer


      answer = MsgBox("Quer Atualizar o Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

      If (answer = vbYes) Then
         data("sinosteel_doc_number") = sinosteel_number_txt.Value
         data("doc_number") = doc_number_txt.Value
         data("name") = doc_name_txt.Value
         data("description") = doc_description_txt.Value
         data("doc_extension") = doc_extension_txt.Value
         data("doc_type_code") = doc_code_txt.Value
         data("doc_format") = doc_formt_txt.Value
         data("pages") = doc_pages_txt.Value
         data("obs") = obs_txt.Value

         where = "id='" & doc_list.Value & "'"
         Call db_documents.update(data, where)

         Call Alert.Show("Dados Atualizados com Sucesso!!!", "", 2000)
      End If

   Else

      MsgBox "Favor completar os daods", , "Dados incompletos"

   End If


End Function



Private Function validate_form() As Boolean

   If (sinosteel_number_txt.Value <> "" And doc_number_txt.Value <> "" And doc_name_txt.Value <> "" And doc_description_txt.Value <> "" And doc_extension_txt.Value <> "") Then
      validate_form = True
   Else
      validate_form = False
   End If
End Function

Private Sub UserForm_Activate()
   Call Shared_DocCategorySelectComp.Mount(doc_category_select)
   getProjectPropertiesHandler
   getDisciplineHandler
   getProjectEquipaments

End Sub

Private Function getProjectPropertiesHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = database_doc_proprty_types.getAll()

   doc_property_select.Clear
   Call Shared_CommonSelectComp.Mount(doc_property_select, respQuery)
End Function

Private Function getDisciplineHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = database_discipline.getAll()

   discipline_select.Clear
   Call Shared_CommonSelectComp.Mount(discipline_select, respQuery)
End Function

Private Function getProjectEquipaments()
   Dim respQuery As ADODB.Recordset
   Set respQuery = database_equipaments.getAll()

   equipament_select.Clear
   Call Shared_CommonSelectComp.Mount(equipament_select, respQuery)
End Function



Private Function load_doc_equipament_handler()
   If (doc_list.Value <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_documents_equipaments.getAll(doc_list.Value)

      equipament_list.Clear



      Do Until respQuery.EOF

         equipament_list.AddItem XdbFactory.getData(respQuery, "id")

         equipament_list.List(equipament_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")
         equipament_list.List(equipament_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "type")


         respQuery.MoveNext
      Loop

   End If
End Function


Private Function load_doc_props_handler()
   If (doc_list.Value <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_document_props.getAll(doc_list.Value)

      doc_properties_list.Clear



      Do Until respQuery.EOF

         doc_properties_list.AddItem XdbFactory.getData(respQuery, "id")

         doc_properties_list.List(doc_properties_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")
         doc_properties_list.List(doc_properties_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "value")


         respQuery.MoveNext
      Loop

   End If
End Function


Private Function get_doc_categorie_handler()
   If (doc_list.Value <> "") Then
      Dim docQuery As ADODB.Recordset
      Set docQuery = db_documents.getDocumentById(doc_list.Value)


      selected_doc_cat_id = XdbFactory.getData(docQuery, "category_id")
      Dim respQuery As ADODB.Recordset

      Set respQuery = db_document_category.getAll(selected_doc_cat_id)

      cat_selected_lb.Caption = XdbFactory.getData(respQuery, "name")

   End If
End Function



Private Function get_doc_discipline_handler()
   If (doc_list.Value <> "") Then
      Dim docQuery As ADODB.Recordset
      Set docQuery = db_documents.getDocumentById(doc_list.Value)


      selected_doc_cat_id = XdbFactory.getData(docQuery, "discipline_id")
      Dim respQuery As ADODB.Recordset

      Set respQuery = db_documents_discipline.getAll(selected_doc_cat_id)

      discipline_selected_lb.Caption = XdbFactory.getData(respQuery, "name")

   End If
End Function
