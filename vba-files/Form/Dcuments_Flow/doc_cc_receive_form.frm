VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_cc_receive_form 
   Caption         =   "Recebimento de Documentos Comentados"
   ClientHeight    =   12015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19635
   OleObjectBlob   =   "doc_cc_receive_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_cc_receive_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Dcuments_Flow


Private project_selected_id
Dim files_dict As Object
Dim import_files_folder_path As String
Dim doc_review_id_selected As String









Private Sub UserForm_Initialize()
   GetDocStatusTypesHandler
   GetDocIssueTypesHandler
   receive_date.Value = Date
End Sub

Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub


Private Function GetDocIssueTypesHandler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getDocIssueTypes()
   select_next_issue.Clear
   Do Until respQuery.EOF

      select_next_issue.AddItem XdbFactory.getData(respQuery, "tag")
      select_next_issue.List(select_next_issue.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

      respQuery.MoveNext

   Loop

End Function




Private Sub change_status_btn_Click()

   change_status_handler

End Sub


Private Function change_status_handler()
   change_status = False
   If (status_select.Value <> "") Then
      For i = 0 To doc_list.ListCount - 1
         If doc_list.Selected(i) = True Then
            doc_list.List(i, 4) = Xhelper.iff(next_rev_txt.Value <> "", next_rev_txt.Value, doc_list.List(i, 3))
            doc_list.List(i, 5) = Xhelper.iff(select_next_issue.Value <> "", select_next_issue.Value, doc_list.List(i, 4))
            doc_list.List(i, 6) = Xhelper.iff(status_select.Value <> "", status_select.Value, doc_list.List(i, 6))

            change_status = True
         End If
      Next i

      If (change_status) Then
         Call Alert.Show("Status do Documento Modificado!!!", "", 2500)
      End If
   End If
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


Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")

      docs_import_fr.Enabled = True
      set_issue_fr.Enabled = True
      read_files_btn.Enabled = True

   Else

      docs_import_fr.Enabled = False
      set_issue_fr.Enabled = False
      read_files_btn.Enabled = False

   End If
End Sub



Private Sub read_files_btn_Click()

   import_files_folder_path = file_helper.open_folder_dialog
   If (import_files_folder_path <> "") Then
      Call import_comment_files(import_files_folder_path)
      btn_update_files.Enabled = True
   End If


End Sub


Private Sub btn_update_files_Click()
   If (import_files_folder_path <> "") Then
      Call import_comment_files(import_files_folder_path)
   End If
End Sub


Private Sub save_review_btn_Click()
 
   Dim Outlook_App As Object
   On Error Resume Next
   Set Outlook_App = GetObject(, "Outlook.Application")

   If Err.Number = 429 Then
      MsgBox "Erro: Gentileza abrir o Outlook Primeiro", vbCritical
   Else


      If (MsgBox("Tem certeza que quer cadastrar o Status do(s) Documento(s)?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")) Then

         If (is_form_valid) Then

            If (copyFilesHandler) Then
               save_status_handler
               Call send_notifi_handler(Outlook_App)

               doc_list.Clear
               Call Alert.Show("Status do docuemnto salvo com Sucesso", "", 2500)
            End If
         Else

            Call Alert.Show("Erro: Favor preencher todos os dados corretamente", "", 2500)
         End If


      End If

   End If
End Sub

Private Function import_comment_files(folder_path As String)


   has_document = False


   Set files_dict = CreateObject("Scripting.Dictionary")

   Dim file_info As Object
   Set file_info = CreateObject("Scripting.Dictionary")

   Dim doc_id As String

   doc_list.Clear
   Set files_dict = file_helper.get_files_from_folders2(folder_path)

   Dim doc_code As String
   Dim doc_name As String
   Dim file_not_found As Integer
   Dim files_found_on_db As Integer
   Dim queryData As Object
   Set queryData = CreateObject("Scripting.Dictionary")
   files_found_on_db = 0
   file_not_found = 0

   Load doc_not_found_form
   doc_not_found_form.doc_not_found_list.Clear

   For Each varKey In files_dict.Keys()
      If (varKey <> "count") Then
         file_name = files_dict(varKey)("file")

         file = Split(UCase(file_name), " - ")
         extension = files_dict(varKey)("extension")


         queryData("PROP1") = project_selected_id
         queryData("PROP2") = "doc_number" 'Search Field
         queryData("PROP3") = Trim(UCase(file(0)))  'Search text (doc number)

         Dim respQuery As ADODB.Recordset
         Set respQuery = XdbFactory.SelectX("get_docs_with_post_status", queryData)


         doc_code = XdbFactory.getData(respQuery, "doc_number")
         rev_id = XdbFactory.getData(respQuery, "rev_id")
         doc_name = XdbFactory.getData(respQuery, "name")
         issue = XdbFactory.getData(respQuery, "issue")
         last_rev = XdbFactory.getData(respQuery, "last_rev")
         doc_extension = XdbFactory.getData(respQuery, "doc_extension")


         If (doc_code <> "") Then

            doc_list.AddItem rev_id
            doc_list.List(doc_list.ListCount - 1, 1) = UCase(doc_code)
            doc_list.List(doc_list.ListCount - 1, 2) = last_rev
            doc_list.List(doc_list.ListCount - 1, 3) = issue
            doc_list.List(doc_list.ListCount - 1, 4) = "PEND"  'Next Review
            doc_list.List(doc_list.ListCount - 1, 5) = "PEND"  'Next Issue
            doc_list.List(doc_list.ListCount - 1, 6) = "PEND"  'status
            doc_list.List(doc_list.ListCount - 1, 7) = doc_name & " - " & XdbFactory.getData(respQuery, "category") & " -> [ " & doc_extension & " ] ( " & extension & " )"
            doc_list.List(doc_list.ListCount - 1, 9) = files_dict(varKey)("path")



            has_document = True

            files_found_on_db = files_found_on_db + 1
         Else
            Call Alert.Show("Documento não Cadastrado no Sistema", doc_code, 2000)
            doc_not_found_form.doc_not_found_list.AddItem doc_code
            file_not_found = file_not_found + 1

         End If


      End If
   Next '

   total_files_found_db_lb.Caption = files_found_on_db

   doc_not_founf_lb.Caption = file_not_found

   total_files_on_folder_lb.Caption = files_dict("count")

   Dim header_titles As Variant

   header_titles = Array("id", "Documento", "Rev.", "TE", "Prox. Rev", "Prox. TE", "Status", "Doc Info")

   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, doc_list_header)

   save_review_btn.Enabled = has_document


End Function





Private Function send_notifi_handler(Outlook_App As Object)


   Dim documents As Object
   Set documents = getDocumentsHandler

   Call act_comments_notifi.make(project_txt.Value, documents, Outlook_App)


End Function


Private Function getDocumentsHandler() As Object

   Dim documents As Object
   Set documents = CreateObject("Scripting.Dictionary")
   For i = 0 To doc_list.ListCount - 1

      Dim document As Object
      Set document = CreateObject("Scripting.Dictionary")

      document("DOC") = doc_list.List(i, 1)
      document("REVIEW") = doc_list.List(i, 2)
      document("ISSUE") = doc_list.List(i, 3)
      document("NEXT_REVIEW") = doc_list.List(i, 4)
      document("NEXT_ISSUE") = doc_list.List(i, 5)
      document("NEXT_STATUS") = doc_list.List(i, 6)
      document("DOC_INFO") = doc_list.List(i, 7)

      documents.Add "n" & i, document
   Next i

   Set getDocumentsHandler = documents
End Function

Private Function save_status_handler()

   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Dim rev_id As String
   Dim next_status As String
   Dim next_review As String
   Dim next_issue As String


   For i = 0 To doc_list.ListCount - 1


      next_review = doc_list.List(i, 4)
      next_issue = doc_list.List(i, 5)
      next_status = doc_list.List(i, 6)


      rev_id = doc_list.List(i, 0)

      data("status") = next_status
      data("next_review") = next_review
      data("next_issue") = next_issue
      data("grd_status") = Trim(UCase(grd_txt.Value))
      data("grd_status_date") = DateHelpers.FormatDateToSQlite(receive_date.Value)
      data("status_date") = DateHelpers.FormatDateToSQlite(Date)

      Dim where As String
      where = "id = " & rev_id

      Call db_documents.updateStatus(data, where)



   Next i



End Function




Private Function is_form_valid() As Boolean


   Dim rev_id As String
   Dim next_status As String
   Dim next_review As String
   Dim next_issue As String
   is_form_valid = False

   For i = 0 To doc_list.ListCount - 1

      next_review = doc_list.List(i, 4)
      next_issue = doc_list.List(i, 5)
      next_status = doc_list.List(i, 6)

      rev_id = doc_list.List(i, 0)

      If (rev_id = "" Or next_status = "PEND" Or next_review = "PEND" Or next_issue = "PEND") Then

         is_form_valid = False

         Exit Function

      End If

   Next i

   is_form_valid = True

End Function


Private Function copyFilesHandler() As Boolean

   Dim review As String
   Dim status As String
   Dim doc_number As String
   Dim destinyFolderPath As String
   Dim new_file_full_path As String
   Dim file_path As String



   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")

   Dim commentsFolder As String
   commentsFolder = Constants.COMMENTS_TEMP_FOLDER & format(Now(), "_dd_MM_yyyy")

   Dim commentsPath As String
   commentsPath = config_sheet.Range("CONFIG_TEMP_FOLDER_PATTH").Value

   Dim commentsFullPath As String
   commentsFullPath = fso.BuildPath(commentsPath, commentsFolder) & "\"

   For i = 0 To doc_list.ListCount - 1

      doc_rev_id = doc_list.List(i, 0)
      doc_number = doc_list.List(i, 1)
      review = doc_list.List(i, 2)
      status = doc_list.List(i, 6)
      document_not_found = False


      For Each varKey In files_dict.Keys()
         If (varKey <> "count") Then
            file_name = Trim(UCase(files_dict(varKey)("file")))
            file_path = files_dict(varKey)("path")
            extension = files_dict(varKey)("extension")

            If Not InStr(file_name, Trim(UCase(doc_number))) = 0 Then

               destinyFolderPath = helper_folder_maker.get_eng_doc_folder(project_selected_id, doc_rev_id)

               new_file_name = doc_number & "_Rev_" & review & "_" & status & "." & extension
               new_file_full_path = destinyFolderPath & "\" & new_file_name

               'Copy file to App Folder
               If (file_helper.copyFilesWithCheckSum(file_path, commentsFullPath & new_file_name)) Then
                  'Move File to user tmp folder
                  Call file_helper.moveFilesWithCheckSum(file_path, new_file_full_path)
                  document_not_found = True
               Else
                  copyFilesHandler = False
                  Exit Function
               End If


            End If

         Else

         End If

      Next 'next

   Next i
   copyFilesHandler = True

End Function


Private Sub doc_list_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   Dim doc_file As String
   Dim doc_full_file_path As String
   Dim temp_file_path As String

   doc_file = doc_list.List(doc_list.ListIndex, 9)

   If (doc_file <> "") Then

      temp_file_path = action_file.copy_to_temp_folder(doc_file)
      Call file_helper.open_file(temp_file_path)

   End If
End Sub


Private Sub doc_list_Change()
   populate_doc_review_list
End Sub




Private Function populate_doc_review_list()

   Dim docQuery As ADODB.Recordset
   Dim reviewQuery As ADODB.Recordset


   For i = 0 To doc_list.ListCount - 1

      If doc_list.Selected(i) = True Then

         doc_review_id_selected = doc_list.List(i, 0)
         If (doc_review_id_selected <> "") Then

            Set docQuery = db_documents.get_doc_by_review2(doc_review_id_selected)
            Set reviewQuery = db_documents.getDocumentReviews(XdbFactory.getData(docQuery, "id"))

            Call Shared_CommonSelectComp.Mount(reviews_select, reviewQuery, "id", "rev_code")
         End If



         Exit Function
      End If
   Next i


End Function

Private Sub change_review_btn_Click()



   If (reviews_select.Value <> "") Then
      For i = 0 To doc_list.ListCount - 1

         If doc_list.Selected(i) = True Then



            doc_list.List(i, 0) = reviews_select.Value
            doc_list.List(i, 2) = reviews_select.List(reviews_select.ListIndex, 1)

            Call Alert.Show("Revisão Modificada com Sucesso!!!", "", 1700)
            Exit Sub
         End If
      Next i

   End If

End Sub

Private Sub btn_open_folder_Click()
   Call file_helper.open_folder(import_files_folder_path & "\")
End Sub



Private Sub remove_doc_btn_Click()
   For i = 0 To doc_list.ListCount - 1
      On Error Resume Next
      If doc_list.Selected(i) = True Then
         doc_list.RemoveItem (i)
      End If
   Next i
End Sub
