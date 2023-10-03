VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grd_recipient_admin_form 
   Caption         =   "Administrar Destinátarios da GRDs"
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21105
   OleObjectBlob   =   "grd_recipient_admin_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grd_recipient_admin_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\GRD

Private selectedProjectId As String
Private search_type As String
Private recipient_id_selected As String
Private recipient_email_id_selected As String
Private emai_msg_conf_id_selected  As String

Private Sub UserForm_Initialize()
   mail_to_select.AddItem "TO"
   mail_to_select.AddItem "CC"
End Sub


Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub



Private Sub add_email_Click()

   can_user_authorized
   add_email_handler
   Call get_all_recipents_emails_handler(recipient_id_selected)
   Call Alert.Show("Email cadastrdo com sucesso!!!", "", 2500)

End Sub


Private Function add_email_handler()

   Dim email As String
   Dim name As String

   name = Trim(UCase(email_user_txt.Value))
   email = Trim(LCase(email_txt.Value))

   If (name <> "" And email <> "") Then
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")
      data("recipient_id") = recipient_id_selected
      data("name") = name
      data("email") = email
      data("type") = mail_to_select.Value



      Call db_grd_recipient.add_email(data)
   Else
      Call Alert.Show("Favor preencher os dados corretamente", "", 2500)
   End If
End Function


Private Sub del_recipient_email_btn_Click()

   can_user_authorized

   Dim answer As Integer
   answer = MsgBox("Quer apagar o email do destinátario da GRD?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
      Dim respQuery As ADODB.Recordset
      Call db_grd_recipient.delete_recipient_email(recipient_email_id_selected, recipient_id_selected)
      Call Alert.Show("E-mail Apagado com Sucesso", "", 2000)
      Call get_all_recipents_emails_handler(recipient_id_selected)

   End If

End Sub
Private Sub del_recipient_btn_Click()

   can_user_authorized

   Dim answer As Integer
   answer = MsgBox("Quer apagar o destinátario de GRD?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_grd_recipient.delete_recipient(recipient_id_selected, selectedProjectId)
      Call Alert.Show("Destinatário Apagado com Sucesso", "", 2000)
      search_select

   End If

End Sub

Private Function can_user_authorized()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd_recipient.get_by_id(recipient_id_selected, selectedProjectId)
   user_id_created_recipient = XdbFactory.getData(respQuery, "user_id")
   If Not (auth.get_user_id = user_id_created_recipient Or auth.is_authorized("SUPER_ADMIN")) Then
      Call Alert.Show("Você não tem Permissão para efetuar esta operação", "", 2000)
      End 'Terminate
   End If
End Function



Private Sub email_list_Change()

   For i = 0 To email_list.ListCount - 1
      If email_list.Selected(i) = True Then
         recipient_email_id_selected = email_list.List(i, 0)
         Exit Sub
      End If
   Next i

End Sub

Private Sub list_all_recipents_btn_Click()


   search_type = "GET_ALL_RECIPIENTS"
   search_select

End Sub


Private Function search_select()

   Select Case search_type
    Case "GET_ALL_RECIPIENTS"
      get_all_recipents_handler
    Case Else

   End Select

End Function


Private Function get_all_recipents_handler()
   If (selectedProjectId <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_grd_recipient.get_all_recipent_from_project(selectedProjectId)
      Call populate_recipient_list(respQuery)
      Call populate_emails_types(selectedProjectId, "1")
   Else
      Call Alert.Show("SELECIONE UM PROJETO PRIMEIRO", "", 2000)
   End If
End Function

Private Function populate_emails_types(project_id As String, section_id As String)

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_email.get_config_msgs(project_id, section_id)
   Call Shared_CommonSelectComp.Mount(email_type_select, respQuery, "name", "")


End Function

Private Function populate_recipient_list(respQuery As ADODB.Recordset)

   recipient_list.Clear

   Do Until respQuery.EOF

      recipient_list.AddItem XdbFactory.getData(respQuery, "id")
      recipient_list.List(recipient_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")
      recipient_list.List(recipient_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "code")
      recipient_list.List(recipient_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "folder_name")
      recipient_list.List(recipient_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "create_ate")


      respQuery.MoveNext
   Loop


   Dim header_titles As Variant

   header_titles = Array("ID", "Destinatário", "GRD Código", "Pasta da GRD", "Criado em")

   Call Xform.SetColumnWidthsAndHeader(recipient_list, lblHidden, header_titles, recipient_header_list)


End Function


Private Sub recipient_list_Change()


   For i = 0 To recipient_list.ListCount - 1

      If recipient_list.Selected(i) = True Then
         recipient_id_selected = recipient_list.List(i, 0)
         If (recipient_id_selected <> "") Then
            Call get_all_recipents_emails_handler(recipient_id_selected)
            Call get_email_mgs_conf(recipient_id_selected)
            emails_fr.Enabled = True
            update_recipient_fr.Enabled = True
            Exit Sub
         End If
      End If
   Next i
   emails_fr.Enabled = False
   update_recipient_fr.Enabled = False
End Sub

Private Function get_email_mgs_conf(recipient_id_selected As String)

   If (recipient_id_selected <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_email.get_email_conf_from_recipient(recipient_id_selected)
      emai_msg_conf_id_selected = XdbFactory.getData(respQuery, "id")
      actual_email_txt.Value = XdbFactory.getData(respQuery, "name")

   End If

End Function

Private Function get_all_recipents_emails_handler(recipient_id As String)

   If (recipient_id <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_grd_recipient.get_all_recipent_emails(recipient_id)
      Call populate_recipient_email_list(respQuery)
   Else
      Call Alert.Show("SELECIONE UM PROJETO PRIMEIRO", "", 2000)
   End If


End Function


Private Function populate_recipient_email_list(respQuery As ADODB.Recordset)

   email_list.Clear

   Do Until respQuery.EOF

      email_list.AddItem XdbFactory.getData(respQuery, "id")
      email_list.List(email_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")
      email_list.List(email_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "email")
      email_list.List(email_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "type")
      email_list.List(email_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "create_ate")


      respQuery.MoveNext
   Loop


   Dim header_titles As Variant

   header_titles = Array("ID", "Nome", "e-mail", "Tipo", "Criado em")

   Call Xform.SetColumnWidthsAndHeader(email_list, lblHidden, header_titles, email_header_list)


End Function



Private Sub update_recipient_btn_Click()

   can_user_authorized

   Dim answer As Integer
   answer = MsgBox("Quer Modificar o destinátario de GRD?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
      Call update_grd_doc_selected("name", name_txt.Value, "Nome do Destinatário")
      Call update_grd_doc_selected("code", folder_name_txt.Value, "Código da GRD")
      Call update_grd_doc_selected("folder_name", code_txt.Value, "Nome da Pasta de GRD")
      search_select
   End If



End Sub




Private Sub search_btn_Click()

   recipient_list.Clear
   email_list.Clear


   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")


   End If
End Sub






Private Function update_prop_handler(prop As String, prop_value As Variant, change_type As String)


   If (Not IsNull(prop_value)) Then
      If (prop <> "" And prop_value <> "") Then


         UserFormAlert.Label1.Caption = "Atualizando!!"
         UserFormAlert.Show
         UserFormAlert.Repaint

         Dim data As Object
         Set data = CreateObject("Scripting.Dictionary")
         data(prop) = prop_value


         Dim where As String

         where = "id='" & recipient_id_selected & "' AND  project_id='" & selectedProjectId & "'"
         Call db_grd_recipient.update(data, where)
         changed = True


         Unload UserFormAlert

         If (changed) Then
            Call Alert.Show("Modificados com Sucesso!!!", "[ " & change_type & " ]", 2500)

         End If
      End If
   End If
End Function
