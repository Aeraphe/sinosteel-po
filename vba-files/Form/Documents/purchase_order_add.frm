VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} purchase_order_add 
   Caption         =   "Castrar Ordem de Compra (Create a Purchase Order - P.O )"
   ClientHeight    =   11805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20115
   OleObjectBlob   =   "purchase_order_add.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "purchase_order_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents

Private selectedProjectId

Private isFormValid



Private Sub UserForm_Activate()

 
End Sub


Private Function auth()
  If (auth.is_logged) Then
     

   Else
      doc_add_form.Hide
 
      Call Alert.Show("Favor Logar no Sistema", "", 3000)

   End If
End Function


Private Sub Bootstrap()

  Call Shared_DocCategorySelectComp.Mount(doc_category_select)
      getProjectPropertiesHandler
      getDisciplineHandler
      getProjectEquipaments
      getDocExtensionstHandler
      getDocFormatstHandler
      getDocCodeTypestHandler
   
      isFormValid = False


End Sub
Private Function getContractItemsHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_projects.get_contract_items(selectedProjectId)

   select_contract_item.Clear
   Call Shared_CommonSelectComp.Mount(select_contract_item, respQuery)
End Function


Private Function getDocCodeTypestHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllDocCodes()

   doc_code_seletc.Clear
   Call Shared_CommonSelectComp.Mount(doc_code_seletc, respQuery)
End Function

Private Function getDocFormatstHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllDocFormats()

   doc_format_select.Clear
   Call Shared_CommonSelectComp.Mount(doc_format_select, respQuery)
End Function


Private Function getDocExtensionstHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllExtensions()

   extension_select.Clear
   Call Shared_CommonSelectComp.Mount(extension_select, respQuery)
End Function


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




Private Sub search_project_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")
      Call set_frames_status(True)
      getContractItemsHandler
      Else
      Call set_frames_status(False)
   End If
End Sub


Private Function set_frames_status(ByVal status As Boolean)

doc_info_fr.Enabled = status
doc_prop_fr.Enabled = status
doc_opt_prop_fr.Enabled = status
doc_equipament_fr.Enabled = status

optiona_fr.Enabled = status
add_doc_btn.Enabled = status

End Function

Private Sub add_doc_property_btn_Click()

   If (doc_property_select.Value <> "" And doc_property_value.Value <> "") Then

      doc_properties_list.AddItem doc_property_select.Value
      doc_properties_list.List(doc_properties_list.ListCount - 1, 1) = doc_property_select.List(doc_property_select.ListIndex, 1)
      doc_properties_list.List(doc_properties_list.ListCount - 1, 2) = doc_property_value.Value


   End If
End Sub


Private Sub delete_doc_property_btn_Click()
   On Error Resume Next
   doc_properties_list.RemoveItem (doc_properties_list.ListIndex)
End Sub




Private Sub add_equipament_btn_Click()

   If (equipament_select.Value <> "") Then

      equipament_list.AddItem equipament_select.Value
      equipament_list.List(equipament_list.ListCount - 1, 1) = equipament_select.List(equipament_select.ListIndex, 1)



   End If
End Sub


Private Sub delete_equipament_btn_Click()
   On Error Resume Next
   equipament_list.RemoveItem (equipament_list.ListIndex)
End Sub



Private Sub add_doc_btn_Click()
   addDocumentHandler
End Sub

Private Function addDocumentHandler()



   Dim answer As Integer

   answer = MsgBox("Quer inserir o documento na LD?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes And selectedProjectId <> "") Then
      Dim doc As Object
      Set doc = CreateObject("Scripting.Dictionary")

      If (ValidateForm) Then

         doc("project_id") = selectedProjectId
         doc("doc_number") = Application.WorksheetFunction.Trim(UCase(doc_number_txt.Value))
         doc("sinosteel_doc_number") = Application.WorksheetFunction.Trim(UCase(sinosteel_number_txt.Value))
         doc("name") = Trim(UCase(doc_name_txt.Value))
         doc("description") = Trim(UCase(doc_description_txt.Value))
         doc("category_id") = doc_category_select.Value
         doc("discipline_id") = discipline_select.Value
         doc("doc_type_code") = UCase(doc_code_seletc.List(doc_code_seletc.ListIndex, 1))
         doc("pages") = doc_total_pges_txt.Value
         doc("doc_extension") = UCase(extension_select.List(extension_select.ListIndex, 1))
         doc("doc_format") = UCase(doc_format_select.List(doc_format_select.ListIndex, 1))
         doc("contract_item") = UCase(select_contract_item.List(select_contract_item.ListIndex, 1))
         doc("project_contract_item_id") = select_contract_item.Value

         doc("obs") = UCase(obs_txt.Value)

         Call db_documents.Create(doc, doc_properties_list, equipament_list)
      End If
      ClearFormHandler
   End If
End Function

Private Function ClearFormHandler()
   sinosteel_number_txt.Value = ""
   doc_number_txt.Value = ""
   doc_name_txt.Value = ""
   doc_property_value = ""
   doc_properties_list.Clear
End Function

Private Function ValidateForm() As Boolean

   If (sinosteel_number_txt.Value <> "" And doc_number_txt.Value <> "" And doc_name_txt.Value <> "" And selectedProjectId <> "") Then
      ValidateForm = True
      Exit Function
   End If

   ValidateForm = False
End Function


Private Sub sinosteel_number_txt_Change()
   If (selectedProjectId <> "") Then
      isFormValid = Not documentHasASinosteelNumber
      If (isFormValid) Then
         Label12.Caption = "Valid"
      Else
         Label12.Caption = "Invalid"
      End If

   End If
End Sub


Private Sub doc_number_txt_Change()
   If (selectedProjectId <> "") Then
      isFormValid = Not documentHasACustomerNumber
      If (isFormValid) Then
         Label15.Caption = "Valid"
      Else
         Label15.Caption = "Invalid"
      End If
   End If
End Sub


Private Function documentHasASinosteelNumber() As Boolean

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.SearchInColumn(selectedProjectId, sinosteel_number_txt.Value, "sinosteel_doc_number")


   Do Until respQuery.EOF
      respQuery.MoveNext

      documentHasASinosteelNumber = True

      Exit Function
   Loop


   documentHasASinosteelNumber = False
End Function


Private Function documentHasACustomerNumber() As Boolean

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.SearchInColumn(selectedProjectId, doc_number_txt.Value, "doc_number")


   Do Until respQuery.EOF
      respQuery.MoveNext

      documentHasACustomerNumber = True

      Exit Function
   Loop


   documentHasACustomerNumber = False
End Function
