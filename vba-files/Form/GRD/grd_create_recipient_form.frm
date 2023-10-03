VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grd_create_recipient_form 
   Caption         =   "Criar destinatário de GRD"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15405
   OleObjectBlob   =   "grd_create_recipient_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grd_create_recipient_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\GRD



Private selectedProjectId
Private selectedSupplierId

Private Sub UserForm_Activate()

   If (auth.is_logged) Then



   Else
      Me.Hide

      Call Alert.Show("Favor Logar no Sistema", "", 3000)

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
   End If
End Sub


Private Sub search_supplier_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_SuppliersSelectComp.GetSupplierSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      supplier_txt.Value = data("name")
      selectedSupplierId = data("id")
   End If
End Sub



Private Sub create_btn_Click()


   Dim answer As Integer

   answer = MsgBox("Do you want to create the new document recipient", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes And selectedProjectId <> "" And selectedSupplierId <> "" And name_txt.Value <> "") Then
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      data("project_id") = selectedProjectId
      data("supplier_id") = selectedSupplierId
      data("user_id") = auth.get_user_id
      data("name") = UCase(name_txt.Value)
      data("code") = UCase(code_txt.Value)
      data("folder_name") = UCase(folder_name_txt.Value)
      Call db_grd.Create(data)
   End If

End Sub
