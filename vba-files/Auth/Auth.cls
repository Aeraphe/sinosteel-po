VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Auth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'namespace=vba-files\Auth


 Const MESSAGE_AUTHORIZATION_ERROR As String = "Você não tem Autorização para Acessar está Área"

Option Explicit



'Data Base Connection
Public user_name As String
Public user_login As String
Private user_id As String
Public user_email As String
Private user_role As String
Private loggin As Boolean



Private Sub Class_Initialize()
    loggin = False
End Sub

Public Function is_logged() As Boolean

    is_logged = loggin
End Function


Public Function is_logged_to_access(Optional form_activate As Variant)
    If Not is_logged Then
        If Not IsMissing(form_activate) Then
            Unload form_activate
        End If
        Alert.Show "Favor Logar no Sistema", "", 2000
        End 'Terminate all process
    End If
End Function





Public Function is_authorized(access_role As String, Optional form_activate As Variant) As Boolean
    is_authorized = False
    If (user_role <> access_role) Then
        If Not IsMissing(form_activate) Then
            Unload form_activate
        End If
        Call Alert.Show(MESSAGE_AUTHORIZATION_ERROR, "", 2000)
        End
    Else
        is_authorized = True
    End If
End Function



Public Function authenticate(login As String, password As String) As Boolean
    Dim password_en As String
    Dim respQuery As ADODB.Recordset

    password_en = helper_encrypt.EncryptStringTripleDES(password)
    Set respQuery = db_auth.get_user(login, password_en)

    If Not respQuery.EOF Then
        loggin = True
        With respQuery
            user_name = .fields("name").Value
            user_id = .fields("id").Value
            user_login = .fields("login").Value
            user_email = .fields("email").Value
            user_role = .fields("role").Value
        End With
    Else
        loggin = False
    End If

    authenticate = loggin
End Function



Public Function loggout()
    loggin = False

End Function


Public Function get_user_id() As Long

    get_user_id = user_id

End Function


Public Function get_user_role() As String
    If Not loggin Then
        get_user_role = ""
        Exit Function
    End If

    get_user_role = user_role
End Function

