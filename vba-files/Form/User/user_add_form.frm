VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} user_add_form 
   Caption         =   "Cadastrar Usuário"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   OleObjectBlob   =   "user_add_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "user_add_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\User




Private Sub UserForm_Activate()
   Call auth.is_logged_to_access(Me)
   Call auth.is_authorized("SUPER_ADMIN", Me)

   getUserRolesHandler
End Sub



Private Function getUserRolesHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_auth.get_all_roles()

   role_select.Clear
   Call Shared_CommonSelectComp.Mount(role_select, respQuery)
End Function



Private Sub add_btn_Click()

   add_user_handler

End Sub


Private Function add_user_handler()
   Dim answer As Integer

   answer = MsgBox("Quer Cadastrar o Usuário?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer = vbYes) Then
      Dim password As String
      password = helper_encrypt.EncryptStringTripleDES(password_txt.Value)
      Dim user_id As Long
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      data("name") = user_name_txt.Value
      data("login") = user_login.Value
      data("phone") = user_phone_txt.Value
      data("email") = user_email_txt.Value
      data("role_id") = role_select.Value
      data("password") = password

      user_id = db_users.Create(data)
   End If
End Function
