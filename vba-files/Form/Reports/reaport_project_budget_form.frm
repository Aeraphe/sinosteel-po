VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} reaport_project_budget_form 
   Caption         =   "Project Budget Report Generate"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "reaport_project_budget_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "reaport_project_budget_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Reports



Private Sub search_project_btn_Click()
   search_project_form.Show

   Dim projectId  As String
   Dim listRowSelected As Long
   Dim params As Object


   If (search_project_form.projects_list.Value <> "") Then

      Set params = CreateObject("Scripting.Dictionary")
      params("project_id") = search_project_form.projects_list.Value

      listRowSelected = search_project_form.projects_list.ListIndex
      project_txt.Value = search_project_form.projects_list.List(listRowSelected, 1)

      Dim response As ADODB.Recordset
      Set response = Report_ProjectBudget_DataBase.GetProjectBudgets(params)
      
      Call ProjectBudgetReport_View.publish(response)
   End If
   Unload search_project_form

End Sub
