VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} reports_doc_flow_form 
   Caption         =   "Gerar Relatórios de Recebimento de Docuemtnos"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270
   OleObjectBlob   =   "reports_doc_flow_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "reports_doc_flow_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Reports


Private selectedProjectId As String



Private Sub UserForm_Activate()
   date_txt.Value = Date
End Sub

Private Sub search_btn_Click()

   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")

      fr_report_options.Enabled = True
   Else
      fr_report_options.Enabled = False
   End If
End Sub



Private Sub btn_create_report_per_date_Click()

   Me.Hide
   Dim answer As Integer

   answer = MsgBox("Quer Gerar o Relatório(s) selecionado(s)?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes And selectedProjectId <> "") Then

      If (report1_ck And date_txt.Value <> "") Then
         Call view_doc_flow_report.publish(selectedProjectId, date_txt.Value)
      End If

      Call Alert.Show("Relatórios Gerados com Sucesso!!!", "", 2000)
   End If

   Me.Show
End Sub
