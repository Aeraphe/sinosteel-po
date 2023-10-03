VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} add_project_form 
   Caption         =   "Add New Project"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   OleObjectBlob   =   "add_project_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "add_project_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Core



Private Sub add_project_btn_Click()


   Dim answer As Integer

   answer = MsgBox("Do you want to create the Project", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      data("project_contract_code") = project_contract_code_txt.Value
      data("project_code") = project_code_txt.Value
      data("project_ch_code") = porject_china_code_txt.Value
      data("name") = project_name_txt.Value
      data("description") = description_txt.Value
      data("start_date") = start_date_txt.Value
      data("finish_date") = finish_date_txt.Value
      data("budget_hours") = total_budget_txt.Value
      data("budget_price") = total_budget_hours.Value



      Call ProjectDataBase.Create(data)

      Unload Me

      add_project_form.Show
   End If
End Sub
