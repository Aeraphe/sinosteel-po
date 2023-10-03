VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_review_import_form 
   Caption         =   "Emitir documento no CDOC (SInosteel)"
   ClientHeight    =   14505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23505
   OleObjectBlob   =   "doc_review_import_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_review_import_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Dcuments_Flow


Public project_selected_id
Public document_selected_id As String
Public docsRequested As Object
Private files_dict As Object
Private is_reviews_valid As Boolean
Public import_files_folder_path As String
Private document_file_path As String




Private Sub UserForm_Activate()

   If (auth.is_logged) Then

      receive_date.Value = Date
   Else
      Me.Hide

      Call Alert.Show("Favor Logar no Sistema", "", 3000)

   End If
End Sub




Private Sub show_not_fount_btn_Click()
   doc_not_found_form.Show

End Sub



Private Sub show_rev_invalid_btn_Click()
   sh_info_list_form.Show

End Sub



Private Sub read_files_btn_Click()
   import_files_folder_path = file_helper.open_folder_dialog
   readFilesAction

End Sub

Private Sub btn_update_files_Click()
   readFilesAction
End Sub

Public Function readFilesAction()

   If (import_files_folder_path <> "") Then
      Call import_review_files_handler(import_files_folder_path)

      docs_import_fr.Enabled = True
      serach_doc_fr.Enabled = True

      Call is_form_valid
   End If

End Function



Private Function import_review_files_handler(ByVal folder_path As String)



   Dim doc_id As String
   Dim reviewValidation As Object
   Dim has_invalid_review As Boolean
   Dim doc_code As String
   Dim doc_name As String
   Dim file_not_found As Integer
   Dim files_found_on_db As Integer
   Dim header_titles As Variant
   Dim newIssue As String
   Dim docRequestIssueIndex As String



   Set files_dict = CreateObject("Scripting.Dictionary")


   'use this form for list the errors on review
   Load sh_info_list_form
   sh_info_list_form.info_list.Clear


   has_invalid_review = False


   doc_list.Clear
   Set files_dict = file_helper.get_files_from_folders(folder_path)


   files_found_on_db = 0
   file_not_found = 0
   Load doc_not_found_form
   doc_not_found_form.doc_not_found_list.Clear

   For Each varKey In files_dict.Keys()
      If (varKey <> "count" And varKey <> "") Then


         file_name = files_dict(varKey)

         file = Split(UCase(file_name), "_REV_")
         On Error GoTo error_handler
         extension = Split(file(1), ".")

         next_rev = extension(0)


         Dim respQuery As ADODB.Recordset
         Set respQuery = db_documents.SearchLimit(project_selected_id, Trim(UCase(file(0))), "doc_number")
         doc_code = XdbFactory.getData(respQuery, "doc_number")
         doc_id = XdbFactory.getData(respQuery, "id")

         If (doc_code <> "") Then

            doc_name = Left(XdbFactory.getData(respQuery, "name") & " - " & XdbFactory.getData(respQuery, "description"), 80)

            issue = respQuery.fields.item("issue")
            issue = Xhelper.iff(IsNull(issue), "-1", issue)

            last_rev = respQuery.fields.item("last_rev")
            last_rev = Xhelper.iff(IsNull(last_rev), "-1", last_rev)

            Set reviewValidation = helper_review.check_review(doc_id, last_rev, next_rev)

            If (Not reviewValidation("status")) Then
               has_invalid_review = True
               sh_info_list_form.info_list.AddItem doc_code & ":  -  Revisão na LD: [ " & last_rev & " ] ->>  Revisão a ser Emitida ( " & next_rev & " )  --->>  " & reviewValidation("msg")


            End If

            docRequestIssueIndex = "ISSUE-" & doc_id
            newIssue = docsRequested(docRequestIssueIndex)

            doc_list.AddItem doc_id
            doc_list.List(doc_list.ListCount - 1, 1) = UCase(doc_code) & " -->> " & reviewValidation("type")
            doc_list.List(doc_list.ListCount - 1, 2) = last_rev
            doc_list.List(doc_list.ListCount - 1, 3) = next_rev
            doc_list.List(doc_list.ListCount - 1, 4) = "[ " & issue & " ]"
            doc_list.List(doc_list.ListCount - 1, 5) = Xhelper.iff(newIssue <> "", newIssue, -1)
            doc_list.List(doc_list.ListCount - 1, 6) = "[ " & XdbFactory.getData(respQuery, "status") & " ]"
            doc_list.List(doc_list.ListCount - 1, 7) = doc_name & " -->> " & XdbFactory.getData(respQuery, "category")
            doc_list.List(doc_list.ListCount - 1, 8) = XdbFactory.getData(respQuery, "pages")
            doc_list.List(doc_list.ListCount - 1, 9) = file_name

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

   header_titles = Array("ID", "Nº Documentos", "Rev. Atual", "Prox. Rev.", "TE Atual", "Prox. TE", "Status", "Descrição", "Páginas")
   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, doc_list_header)

   If (has_invalid_review) Then

      sh_info_list_form.title_lb.Caption = "Erro: Os Documentos Listados Abaixo Estão Com Revisão INVÁLIDA "
      sh_info_list_form.Repaint
      sh_info_list_form.Show
      total_rev_invalid_lb.Caption = sh_info_list_form.info_list.ListCount
      is_reviews_valid = False

   Else
      total_rev_invalid_lb.Caption = 0
   End If

   Exit Function
error_handler:
   MsgBox "Erro: Documento Fora do Formato: " & varKey

End Function


Private Sub btn_post_and_open_grd_Click()

   Dim docList As Object
   Set data = CreateObject("Scripting.Dictionary")

   If (is_form_valid) Then


      answer = MsgBox("Tem certeza que quer emitir os Docuemntos para o CDOC e Criar A GRD?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

      If (answer = vbYes) Then

         Call Alert.Show("Emitindo os Docuemtnos para o Sistema CDOC", "Liberando o Acesso para a Eng. SENOSTEEL", 2500)
         Set docList = addDocumentReview

         Load grd_create_confirmation_form

         Set grd_create_confirmation_form.docList = CreateObject("Scripting.Dictionary")

         grd_create_confirmation_form.project_txt.Value = project_txt.Value
         grd_create_confirmation_form.projectIdSelected = project_selected_id
         Set grd_create_confirmation_form.docList = docList
         grd_create_confirmation_form.Show
         Unload Me

      End If
   Else

      Call Alert.Show("Erro: Favor Definir as Emissões dos Docuemtnos", "Está faltando definier as Emissões", 2500)


   End If


End Sub

Private Sub save_review_btn_Click()

   ' Check if the form is valid and ask for confirmation before emitting the documents
   If (is_form_valid And MsgBox("Tem Certeza Que quer emitir os documentos para o CDOC SINOSTEEL?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!") = vbYes) Then

      Call addDocumentReview(doc_review_save_confirm_form.options)

   Else

      Call Alert.Show("Erro: Favor Definir as Emissões dos Documentos", "Está faltando definir as Emissões", 2000)

   End If

End Sub





'/*
' This function adds a document review for each document in the `doc_list` object
'*/
Private Function addDocumentReview() As Object

   ' Declare necessary variables
   Dim data As Object ' holds data to be inserted into the database
   Dim file_name As String ' name of the file being added
   Dim doc_id As String ' document ID
   Dim next_rev As String ' next revision code
   Dim doc_number_splited() As String ' array holding the split document number and validation status
   Dim respQueryDocReview As ADODB.Recordset ' response object from searching for an existing document review
   Dim docRequestId As String ' ID of the document request
   Dim docList As Object ' dictionary object to hold documents for sending to GRD
   Dim docListIndex As String ' index used for adding documents to docList
   Dim rev_id As Long ' ID of the document revision

   ' Create dictionary objects for `data` and `docList`
   Set data = CreateObject("Scripting.Dictionary")
   Set docList = CreateObject("Scripting.Dictionary")

   ' Display progress dialog box
   Load UserFormAlert
   UserFormAlert.Label1 = "Cadastrando o(s) Documento(s)" ' Set the label on the progress dialog box.
   UserFormAlert.Show

   ' Create object for accessing the file system
   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")

   On Error GoTo ErrorHandler ' Jump to error handler if there's any error within the code block.

   For i = 0 To doc_list.ListCount - 1

      ' Get relevant values from doc_list
      doc_id = doc_list.List(i, 0)
      next_rev = doc_list.List(i, 3)
      file_name = doc_list.List(i, 9)
If (doc_id <> "" And file_name <> "") Then
      ' Split document number and validation status
      doc_number_splited = Split(doc_list.List(i, 1), "-->>")
      doc_rev_validation = Trim(doc_number_splited(1))

      ' Get document request ID
      docRequestId = getRequstDocIdHandler(doc_id)

      If (doc_rev_validation <> "GRD") Then 'Check if the document is not already a GRD document.

         ' Prepare data for insertion
         data("doc_id") = doc_id
         data("rev_code") = next_rev
         data("issue") = doc_list.List(i, 5)
         data("status") = Constants.REVIEW_SATUS_SEND
         data("grd") = Trim(UCase(grd_txt.Value))
         data("grd_date") = DateHelpers.FormatDateToSQlite(receive_date.Value)
         data("request_doc_id") = docRequestId
         data("file_name") = file_name
         data("file_extension") = fso.GetExtensionName(file_name)

         ' Copy file to project folder
         ' Check if the file was successfully copied
         ' Insert document review into database
         ' Change the document request status to the LIB_ENG stage.
         If (shared_project_files.copyFile(import_files_folder_path & file_name, project_selected_id, doc_id, file_name)) Then
            rev_id = db_documents.InsertDocumentReview(data)
            Call helper_log.DebugApp("InsertDocumentReview: " & file_name & " " & next_rev)

            If rev_id <> 0 And IsNumeric(rev_id) Then
               Call act_doc_request.changeDocRequestStatus(docRequestId, Constants.LIB_ENG)
               Call helper_log.DebugApp("changeDocRequestStatus: " & file_name & " " & next_rev)
            End If
         End If

      Else
         ' If document is already a GRD document, search for existing document review
         Set respQueryDocReview = db_documents.search_doc_by_review(doc_id, next_rev)
         rev_id = XdbFactory.getData(respQueryDocReview, "id")
      End If

      ' Add document to list for sending to GRD
      ' only If the document has a non-zero revision and is numeric.
      ' Add document to dictionary object - docList
      If rev_id <> 0 And IsNumeric(rev_id) Then
         docListIndex = "n" & rev_id
         docList.Add docListIndex, preparDocToSendGrd(rev_id)
         Call helper_log.DebugApp("Add document to list for sending to GRD: " & file_name & " " & next_rev)
      End If


      ' Update progress dialog
      UserFormAlert.labelInfo.Caption = "Documento: " & doc_list.List(i, 1)
      UserFormAlert.Repaint
End If
   Next i

   Unload UserFormAlert ' close progress dialog box

   ' Display success message
   Call Alert.Show("Documentos Cadastrados Com Sucesso", "", 2500)

   ' Set the return value to the `docList` dictionary object
   Set addDocumentReview = docList

   Exit Function ' exit function block and execute any following instructions.

ErrorHandler:
   ' Close progress dialog and show error message
   Unload UserFormAlert
   Call helper_log.DebugApp("Error on Copy Files to Eng Folder: " & Err.Number & ": " & Err.description)

End Function


'/*
'
'
'
'*/
Private Function getRequstDocIdHandler(ByVal docId As String) As String

   Dim requestDocId() As String

   For Each varKey In docsRequested.Keys()
      If (docId = docsRequested(varKey)) Then
         requestDocId = Split(varKey, "-")
         getRequstDocIdHandler = requestDocId(1)
         Exit Function
      End If
   Next varKey
End Function






Private Function preparDocToSendGrd(docRevId As Long) As Object

   doc_name_split = Split(doc_list.List(i, 1), " -->> ")

   doc_cat = Split(doc_list.List(i, 7), " -->> ")


   Dim doc As Object
   Set doc = CreateObject("Scripting.Dictionary")

   doc("docRevId") = docRevId
   doc("docNumber") = doc_name_split(0) 'Doc Number
   doc("docNextRev") = doc_list.List(i, 3) 'Next Rev
   doc("desciption") = doc_list.List(i, 7)  'Description
   doc("category") = doc_cat(1)  'Category
   doc("type") = "OR" 'type
   doc("media") = "MD" 'media
   doc("copies") = 1  'copies
   doc("pages") = doc_list.List(i, 8)  'pages

   Set preparDocToSendGrd = doc

End Function



Private Function is_form_valid() As Boolean

   Dim doc_number_splited() As String
   Dim doc_rev_validation As String
   Dim form_state As Boolean
   Dim total_docs As Long


   form_state = True
   total_docs = doc_list.ListCount

   If (total_docs > 0) Then
      For i = 0 To total_docs - 1

         doc_number_splited = Split(doc_list.List(i, 1), "-->>")
         doc_rev_validation = Trim(doc_number_splited(1))

         If (doc_rev_validation = "REV_VALIDA") Then

            file = Split(UCase(file_name), "_REV_")

            If (doc_list.List(i, 5) = "-1" Or doc_list.List(i, 5) = "" Or doc_list.List(i, 3) = "-1" Or doc_list.List(i, 3) = "") Then
               form_state = False
               Exit For
            End If

         End If

         If (doc_rev_validation = "REV_ERROR") Then

            form_state = False
            Exit For
         End If

      Next i
   Else
      form_state = False
   End If

   save_review_btn.Enabled = form_state

   is_form_valid = form_state
End Function



Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")
   clear_form_fields

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")

      serach_doc_fr.Enabled = True
      Call form_frames_state(True)

   Else


      Call form_frames_state(False)
   End If
End Sub



Private Sub doc_list_Change()
   For i = 0 To doc_list.ListCount - 1
      If doc_list.Selected(i) = True Then
         document_selected_id = doc_list.List(i, 0)
         Call get_selected_doc_info(document_selected_id)
         fr_doc_titles.Enabled = True
         document_file_path = helper_folder_maker.get_eng_doc_folder(project_selected_id, document_selected_id, "SENT")
         Exit Sub
      End If

   Next i
   fr_doc_titles.Enabled = False
End Sub


Private Function get_selected_doc_info(ByVal doc_id As String)
   If (doc_id <> "") Then
      Dim respQuery As ADODB.Recordset
      Set respQuery = db_documents.getDocumentById(doc_id)

      doc_name_txt.Value = XdbFactory.getData(respQuery, "name")
      doc_description_txt.Value = XdbFactory.getData(respQuery, "description")
      doc_number_acivate_txt.Value = XdbFactory.getData(respQuery, "doc_number")


   End If
End Function






Private Sub remove_doc_btn_Click()
   For i = 0 To doc_list.ListCount - 1
      On Error Resume Next
      If doc_list.Selected(i) = True Then
         doc_list.RemoveItem (i)
      End If
   Next i
   is_form_valid
End Sub



Private Sub doc_list_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


   Dim doc_file As String
   Dim doc_full_file_path As String
   Dim temp_file_path As String


   doc_file = doc_list.List(doc_list.ListIndex, 9)

   If (doc_file <> "" And import_files_folder_path <> "") Then

      doc_full_file_path = import_files_folder_path & "\" & doc_file
      temp_file_path = action_file.copy_to_temp_folder(doc_full_file_path)
      Call file_helper.open_file(temp_file_path)

   End If

End Sub

Private Sub btn_open_folder_Click()
   Call file_helper.open_folder(document_file_path & "\")
End Sub


Private Sub btn_add_grd_Click()

   Dim answer As Long
   Dim doc_number_splited() As String
   Dim doc_rev_validation As String

   answer = MsgBox("Tem certeza que o documento já foi emitido?" & vbNewLine & " O Documento só será considerado na GRD", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then


      Load UserFormAlert
      UserFormAlert.Label1 = "Incluíndo Documentos só na GRD"
      Call Xhelper.waitMs(2000)

      For i = 0 To doc_list.ListCount - 1
         On Error Resume Next
         If doc_list.Selected(i) = True Then
            doc_number_splited = Split(doc_list.List(i, 1), "-->>")
            doc_rev_validation = Trim(doc_number_splited(0))
            doc_list.List(i, 1) = doc_rev_validation & "-->>GRD"
            UserFormAlert.labelInfo.Caption = "Documento: " & doc_rev_validation
            UserFormAlert.Repaint
         End If
      Next i
      is_form_valid


      UserFormAlert.Label1 = "Finalizado"
      UserFormAlert.labelInfo.Caption = ""
      Call Xhelper.waitMs(2000)
   End If
End Sub
