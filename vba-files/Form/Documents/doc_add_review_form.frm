VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_add_review_form 
   Caption         =   "Subir Revis�o do Docuemtno"
   ClientHeight    =   11640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15615
   OleObjectBlob   =   "doc_add_review_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_add_review_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents





Private selectedProjectId
Private selectedSupplierId
Private isFormValid
Private flow_id



Private Sub clear_btn_Click()
   doc_flow_name_txt.Enabled = True
   doc_flow_name_txt.Value = ""
   flow_id = ""
End Sub



Private Sub UserForm_Activate()

   If (auth.is_logged) Then
      GetDocIssueTypesHandler
      GetDocStatusTypesHandler


   Else
      Me.Hide

      Call Alert.Show("Favor Logar no Sistema", "", 3000)

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

Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")

   End If
End Sub



Private Function SearchDocumentHandler(tb As String)


   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.search(selectedProjectId, search_doc_txt.Value, tb)
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
   SearchDocumentRev
   get_next_review
   get_all_doc_reviews
End Sub



Private Function SearchDocumentRev()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.SearchLastRev(doc_list.Value)


   last_rev.Value = ""
   last_grd_date.Value = ""
   last_issue.Value = ""
   last_status.Value = ""
   last_grd_txt.Value = ""

   Do Until respQuery.EOF

      last_rev.Value = XdbFactory.getData(respQuery, "rev_code")
      last_grd_date.Value = format(CDate(XdbFactory.getData(respQuery, "grd_date")), "DD/MM/YYYY")
      last_issue.Value = XdbFactory.getData(respQuery, "issue")
      last_status.Value = XdbFactory.getData(respQuery, "status")
      last_grd_txt.Value = XdbFactory.getData(respQuery, "grd")

      respQuery.MoveNext

   Loop

End Function

Private Function get_next_review()
   Dim last_review As Variant

   Dim next_review As Variant
   If (last_rev.Value = "") Then
      next_review = "A"
   Else
      On Error Resume Next
      last_review = CInt(last_rev.Value)


      If (VarType(last_rev.Value) = vbString And last_review = Empty) Then
         next_review = helper_string.NextLetter(last_rev.Value)
      Else
         next_review = last_review + 1
      End If


   End If

   review_txt.Value = next_review

End Function


Private Sub review_txt_Change()
   If (Trim(review_txt.Value) <> "") Then
      validate_review
   End If
End Sub


Private Function validate_review() As Boolean



   If (last_rev.Value = "") Then
      If (review_txt.Value = "A" Or review_txt.Value = 0) Then
         validate_review = True
         Exit Function
      End If
   Else

      On Error Resume Next
      last_review = CInt(last_rev.Value)
      On Error Resume Next
      nextf = CInt(review_txt.Value)




      If (VarType(last_rev.Value) = vbString And nextf = Empty) Then
         next_review = helper_string.NextLetter(last_rev.Value)
         If (review_txt.Value = next_review) Then
            validate_review = True
            Exit Function
         End If
      ElseIf (last_review + 1 = CInt(review_txt.Value)) Then
         validate_review = True
         Exit Function
      End If
   End If
   validate_review = False
End Function

Private Function GetDocIssueTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getDocIssueTypes()
   issue_select.Clear
   Do Until respQuery.EOF

      issue_select.AddItem XdbFactory.getData(respQuery, "tag")
      issue_select.List(issue_select.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

      respQuery.MoveNext

   Loop

End Function



Private Function GetDocStatusTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getDocStatusTypes()
   status_select.Clear
   Do Until respQuery.EOF

      status_select.AddItem XdbFactory.getData(respQuery, "tag")
      status_select.List(status_select.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

      respQuery.MoveNext

   Loop

End Function



Public Sub add_supplier_btn_Click()
   doc_supplier_form.Show

End Sub



Private Sub add_review_btn_Click()

   Dim answer As Integer
   Dim review_id As Long

   answer = MsgBox("Quer subir a Revis�o do Documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes And FormValidate) Then


      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      data("doc_id") = doc_list.Value
      data("rev_code") = Application.WorksheetFunction.Trim(UCase(review_txt.Value))
      data("issue") = issue_select.Value
      data("grd") = Application.WorksheetFunction.Trim(UCase(grd_txt.Value))
      data("grd_date") = DateHelpers.FormatDateToSQlite(receive_date.Value)
      data("status") = status_select.Value
      data("obs") = obs.Value

      review_id = db_documents.InsertDocumentReview(data)

      SearchDocumentRev
      Call Alert.Show("Rev. Modificada Com Sucess!!!", "", 2500)
   End If



End Sub


Private Function FormValidate() As Boolean

   Dim check_rev As Boolean
   check_rev = check_if_review_exist
   If (Not check_rev) Then
   If (doc_list.Value <> "" And review_txt.Value <> "" And issue_select.Value <> "" And receive_date.Value <> "" And status_select.Value <> "") Then
      FormValidate = True
   Else
      FormValidate = False
     Call Alert.Show("Erro: Preencha os dados Corretamente", "", 2500)
   End If
   Else
     Call Alert.Show("Erro: Revis�o j� existe", "", 2500)
   End If
End Function

Private Function check_if_review_exist() As Boolean


   For i = 0 To doc_review_list.ListCount - 1
      old_rev = doc_review_list.List(i, 1)

      If (review_txt.Value = old_rev) Then

        
         check_if_review_exist = True
         Exit Function

      End If
   Next i
   check_if_review_exist = False
End Function



Private Sub create_doc_flow_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")


   If (auth.is_logged And doc_flow_name_txt.Value <> "") Then
      Dim answer As Integer


      answer = MsgBox("Quer Criar o Fluxo de Documentos:", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

      If (answer = vbYes) Then

         data("name") = doc_flow_name_txt.Value
         data("user_id") = auth.get_user_id
         flow_id = db_doc_flow.Create(data)
         doc_flow_name_txt.Enabled = False
         Call Alert.Show("Criado com Sucesso!!", "", 1500)
      End If
   Else
      Call Alert.Show("Favor Logar no Sistema", "", 2500)
   End If
End Sub



Private Function get_all_doc_reviews()

   If (doc_list.Value <> "") Then

      Dim respQuery As ADODB.Recordset
      Set respQuery = db_documents.getDocumentReviews(doc_list.Value)
      doc_review_list.Clear
      Do Until respQuery.EOF

         doc_review_list.AddItem XdbFactory.getData(respQuery, "id")

         doc_review_list.List(doc_review_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "rev_code")
         doc_review_list.List(doc_review_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "issue")
         doc_review_list.List(doc_review_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "grd")
         doc_review_list.List(doc_review_list.ListCount - 1, 4) = format(CDate(XdbFactory.getData(respQuery, "grd_date")), "DD/MM/YYYY")
         doc_review_list.List(doc_review_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "status")
         doc_review_list.List(doc_review_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "obs")

         respQuery.MoveNext
      Loop
   End If
End Function
