VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sh_info_list_form 
   Caption         =   "Info!!!"
   ClientHeight    =   11115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17370
   OleObjectBlob   =   "sh_info_list_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sh_info_list_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Shared



Private Sub close_btn_Click()
Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "Botão desabilitado", vbCritical
    End If
     
End Sub
