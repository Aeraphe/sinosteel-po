VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} login_form 
   Caption         =   "Logar no Sistema"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "login_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "login_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Core



Private Sub UserForm_Initialize()

   On Error GoTo defaultValue
   ReadUserConfigurationsFromFile
   ck_stayLoggedIn.Value = CBool(core_user_config.userConfigurations("stayLoggedIn"))

   stayLoggedIn ck_stayLoggedIn.Value

   Exit Sub
defaultValue:
   ck_stayLoggedIn.Value = False
End Sub

Private Sub UserForm_Activate()

   If (auth.is_logged) Then
      Unload Me
      Call Alert.Show("Logado com Sucesso", "Usuário: " & auth.user_name, 2500)
   End If
End Sub

Private Sub ck_stayLoggedIn_Click()
   ' Update the "stay logged in" configuration setting
   core_user_config.UpdateStayLoggedInValue ck_stayLoggedIn.Value

End Sub


Private Sub stayLoggedIn(ByVal action As Boolean)

   Dim loginData As Object

   On Error Resume Next
   If (action) Then

      Set loginData = core_stay_loggedin.GetCredentials()
      Call loggin_handler(loginData("savedLogin"), loginData("savedPassword"))
      
      Else
      login_txt.Value = ""
      password_txt.Value = ""

   End If

End Sub

Private Sub loggin_btn_Click()

   Call loggin_handler(login_txt.Value, password_txt.Value)

   If (ck_stayLoggedIn.Value) Then
      Call core_stay_loggedin.SaveCredentials(login_txt.Value, password_txt.Value)
   End If

End Sub


Private Sub password_txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = vbKeyReturn Then

      Call loggin_handler(login_txt.Value, password_txt.Value)

      If (ck_stayLoggedIn.Value) Then
         Call core_stay_loggedin.SaveCredentials(login_txt.Value, password_txt.Value)
      End If

   End If
End Sub


   
Private Function loggin_handler(ByVal login As String, ByVal password As String)

   If (Not auth.is_logged) Then
      If (login <> "" And password <> "") Then

         Call logginAlertHandler(auth.authenticate(login, password))
      End If

   End If
End Function



Private Function logginAlertHandler(ByVal isLogged As Boolean)
   If (isLogged) Then
      Unload Me
      Call Alert.Show("Logado com Sucesso", "Usuário: " & auth.user_name, 2500)
   Else
      Call Alert.Show("Não foi possível Logar no Sistema", "", 2500)
   End If
    Unload Me
End Function
