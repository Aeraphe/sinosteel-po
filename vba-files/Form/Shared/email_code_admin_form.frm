VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} email_code_admin_form 
   Caption         =   "Código dos emails"
   ClientHeight    =   12780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18720
   OleObjectBlob   =   "email_code_admin_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "email_code_admin_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Shared






Public selectedProjectId As String
Public selected_email_id As String



Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub


Private Sub UserForm_Initialize()
   populate_company_sections_select
End Sub


Private Sub css_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 Call edit_text_handler(css_txt)
End Sub

Private Sub midle_msg_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call edit_text_handler(midle_msg_txt)
End Sub

Private Sub msg_first_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call edit_text_handler(msg_first_txt)
End Sub

Private Sub msg_last_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   Call edit_text_handler(msg_last_txt)
End Sub


Private Function edit_text_handler(text_field As MSForms.TextBox)


Load sh_edit_text_form
sh_edit_text_form.msg_txt.Value = text_field.Value
sh_edit_text_form.Show
text_field.Value = sh_edit_text_form.msg_txt.Value


End Function




Private Function populate_company_sections_select()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_app_sections.getAll()

   select_section.Clear
   Call Shared_CommonSelectComp.Mount(select_section, respQuery)
End Function

Private Sub search_btn_Click()
   get_project_handler
   get_email_config_msg
End Sub

Private Function get_project_handler()

   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")
   End If
End Function


Private Function get_email_config_msg()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_email.get_config_msgs(selectedProjectId, "1")
   Call populate_email_list_handler(respQuery)
End Function

Private Function populate_email_list_handler(respQuery As ADODB.Recordset)

   email_msg_list.Clear
   Do Until respQuery.EOF

      email_msg_list.AddItem XdbFactory.getData(respQuery, "id")
      email_msg_list.List(email_msg_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "pre_title")
      email_msg_list.List(email_msg_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "section_name")
      email_msg_list.List(email_msg_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "description")
      email_msg_list.List(email_msg_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "create_ate")

      respQuery.MoveNext
   Loop


   Dim header_titles As Variant

   header_titles = Array("id", "Pre-Titulo", "Setor", "Descrição", "Criado em:")

   Call Xform.SetColumnWidthsAndHeader(email_msg_list, lblHidden, header_titles, email_msg_header_list)


End Function

Private Sub btn_add_Click()

   Call auth.is_authorized("SUPER_ADMIN")
   Dim answer As Integer

   answer = MsgBox("Quer criar o novo e-mail padrão?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
      add_new_email_msg_handler
   End If
End Sub

Private Function add_new_email_msg_handler()


   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")
   data("project_id") = selectedProjectId
   data("user_id") = auth.get_user_id
   data("header_css_msg") = Trim(css_txt.Value)
   data("midle_msg") = Trim(midle_msg_txt.Value)
   data("pre_title") = Trim(pre_title_txt.Value)
   data("name") = Trim(description_txt.Value)
   data("section_id") = select_section.Value
   data("pre_msg") = Trim(msg_first_txt.Value)
   data("pos_msg") = Trim(msg_last_txt.Value)
   Call db_email.Create(data)
   Call Alert.Show("Mensagem Padrão Criada Com Sucesso", "", 2000)

End Function

Private Sub btn_delete_code_Click()

End Sub

Private Sub email_codes_list_Click()

End Sub

Private Sub select_category_Change()

End Sub

Private Sub select_sector_Change()

End Sub
