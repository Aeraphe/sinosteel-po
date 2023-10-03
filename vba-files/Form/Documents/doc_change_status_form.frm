VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_change_status_form 
   Caption         =   "Modificar a situação do documento (Retorno - Fluxo de Aprovação)"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13995
   OleObjectBlob   =   "doc_change_status_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_change_status_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents



Private selectedProjectId
Private selectedDocumentId


Private Sub UserForm_Activate()
   GetDocStatusTypesHandler
   status_date_txt.Value = Date
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



Private Sub search_doc_btn_Click()
   Call SearchDocumentHandler("name")
End Sub


Private Sub search_doc_sinosteel_btn_Click()
   Call SearchDocumentHandler("sinosteel_doc_number")
End Sub


Private Sub search_doc_supplier_btn_Click()
   Call SearchDocumentHandler("doc_number")
End Sub



Private Function SearchDocumentHandler(tb As String)


   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.search(selectedProjectId, doc_txt.Value, tb)

   doc_list.Clear
   Do Until respQuery.EOF

      doc_list.AddItem XdbFactory.getData(respQuery, "id")

      doc_list.List(doc_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "doc_number")
      doc_list.List(doc_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "name")
      doc_list.List(doc_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "category")
      doc_list.List(doc_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "doc_type")


      respQuery.MoveNext
   Loop

End Function




Private Sub doc_list_Click()

   get_doc_review_handler

End Sub

Private Sub doc_review_select_Change()
   setOldDocInfo
End Sub



Private Function get_doc_review_handler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getDocumentReviews(doc_list.Value)

   doc_review_select.Clear

   Do Until respQuery.EOF

      doc_review_select.AddItem XdbFactory.getData(respQuery, "id")
      doc_review_select.List(doc_review_select.ListCount - 1, 1) = XdbFactory.getData(respQuery, "issue")
      doc_review_select.List(doc_review_select.ListCount - 1, 2) = XdbFactory.getData(respQuery, "status")
      doc_review_select.List(doc_review_select.ListCount - 1, 3) = XdbFactory.getData(respQuery, "obs")
      doc_review_select.List(doc_review_select.ListCount - 1, 4) = XdbFactory.getData(respQuery, "rev_code")
      doc_review_select.List(doc_review_select.ListCount - 1, 5) = XdbFactory.getData(respQuery, "grd_date")
      doc_review_select.List(doc_review_select.ListCount - 1, 6) = XdbFactory.getData(respQuery, "grd")

      respQuery.MoveNext
   Loop
   If (doc_review_select.ListCount > 0) Then
      doc_review_select.ListIndex = 0
   End If
   setOldDocInfo
End Function




Private Function setOldDocInfo()
   If (doc_review_select.ListCount > 0) Then
      issue_txt.Value = doc_review_select.List(doc_review_select.ListIndex, 1)
      doc_status_txt.Value = doc_review_select.List(doc_review_select.ListIndex, 2)
      last_obs_txt.Value = doc_review_select.List(doc_review_select.ListIndex, 3)
      old_data_txt.Value = doc_review_select.List(doc_review_select.ListIndex, 5)
      old_grd_txt.Value = doc_review_select.List(doc_review_select.ListIndex, 6)
   Else

      issue_txt.Value = ""
      doc_status_txt.Value = ""
      last_obs_txt.Value = ""
      old_data_txt.Value = ""
      old_grd_txt.Value = ""
   End If
End Function


Private Function GetDocStatusTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getDocStatusTypes()
   doc_status_select.Clear
   Do Until respQuery.EOF

      doc_status_select.AddItem XdbFactory.getData(respQuery, "tag")
      doc_status_select.List(doc_status_select.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

      respQuery.MoveNext

   Loop

End Function



Private Sub change_status_btn_Click()



   Dim answer As Integer
   Dim review_id As Long
   Dim sql_where As String
   Dim data As Object

   answer = MsgBox("Do you want to Change the Document STATE ?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes And doc_review_select.Value <> "" And doc_status_select.Value <> "" And status_date_txt.Value <> "") Then


      
      Set data = CreateObject("Scripting.Dictionary")

      data("status") = Application.WorksheetFunction.Trim(doc_status_select.Value)
      data("grd_date") = Application.WorksheetFunction.Trim(status_date_txt.Value)
      data("grd") = Application.WorksheetFunction.Trim(grd_txt.Value)
      data("obs") = Application.WorksheetFunction.Trim(obs_txt.Value)

      sql_where = "id = '" & doc_review_select.Value & "'"

      Call db_documents.updateStatus(data, sql_where)

      doc_change_status_list.AddItem doc_review_select.Value

      doc_change_status_list.List(doc_change_status_list.ListCount - 1, 1) = doc_status_txt.Value & "---->" & data("status")
      doc_change_status_list.List(doc_change_status_list.ListCount - 1, 2) = doc_list.List(doc_list.ListIndex, 1)
      doc_change_status_list.List(doc_change_status_list.ListCount - 1, 3) = doc_list.List(doc_list.ListIndex, 2)
      doc_change_status_list.List(doc_change_status_list.ListCount - 1, 4) = doc_list.List(doc_list.ListIndex, 3)
      doc_change_status_list.List(doc_change_status_list.ListCount - 1, 5) = doc_list.List(doc_list.ListIndex, 4)

   End If
End Sub


Private Sub remove_btn_Click()
   On Error Resume Next
   doc_change_status_list.RemoveItem (doc_change_status_list.ListIndex)
End Sub
