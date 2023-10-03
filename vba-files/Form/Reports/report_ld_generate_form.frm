VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} report_ld_generate_form 
   Caption         =   "Gerar a Lista de Documentos do Projeto (LD)"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15405
   OleObjectBlob   =   "report_ld_generate_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "report_ld_generate_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Reports


Private selectedProjectId


Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)
   Call auth.is_authorized("SUPER_ADMIN", Me)


End Sub


Private Sub search_project_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")
   End If
End Sub

Private Sub btn_select_ld_file_Click()

   ld_file_txt.Value = file_helper.open_file_dialog

End Sub


Private Sub btn_generate_ld_Click()

   If (ld_file_txt.Value <> "" And selectedProjectId <> "") Then
    Call view_ld.publish(selectedProjectId, ld_file_txt.Value)
   End If
End Sub
