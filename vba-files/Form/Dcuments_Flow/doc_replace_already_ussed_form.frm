VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_replace_already_ussed_form 
   Caption         =   "Subistituir documento ja emitido no Sistema"
   ClientHeight    =   13830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20880
   OleObjectBlob   =   "doc_replace_already_ussed_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_replace_already_ussed_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Dcuments_Flow


Private project_selected_id
Private document_selected_id As String
Private document_review_selected_id As String
Private files_dict As Object
Private files_folder_path As String


Private Sub UserForm_Initialize()
   replace_date_txt.Value = Date
End Sub

Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub

Private Sub btn_replace_doc_Click()
   move_files
End Sub

Private Sub read_files_btn_Click()

   files_folder_path = ""
   who_received_doc_list.Clear
   
   ' Open the select folder prompt
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = -1 Then ' if OK is pressed
         files_folder_path = .selectedItems(1)
      End If
   End With

   If files_folder_path <> "" Then
      Call load_files_handler(files_folder_path)
   End If



End Sub

Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")


   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")
      read_files_btn.Enabled = True
      docs_import_fr.Enabled = True

   Else
      read_files_btn.Enabled = False
      docs_import_fr.Enabled = False
   End If
End Sub

Private Function load_files_handler(folder_path As String)


   Dim doc_id As String
   Dim doc_code As String
   Dim doc_name As String
   Dim file_not_found As Integer
   Dim files_found_on_db As Integer
   Dim issue As String
   Dim files_dict As Object
   Dim last_rev As String
   Dim rev_id As String


   Set files_dict = CreateObject("Scripting.Dictionary")



   doc_list.Clear
   Set files_dict = file_helper.get_files_from_folders(folder_path)


   files_found_on_db = 0
   file_not_found = 0

   doc_not_found_form.doc_not_found_list.Clear

   For Each varKey In files_dict.Keys()
      If (varKey <> "count") Then


         file_name = files_dict(varKey)

         file = Split(UCase(file_name), "_REV_")
         On Error GoTo error_handler
         extension = Split(file(1), ".")

         next_rev = extension(0)


         Dim respQuery As ADODB.Recordset
         Set respQuery = db_documents.SearchLimit(project_selected_id, Trim(UCase(file(0))), "doc_number")
         
         doc_code = XdbFactory.getData(respQuery, "doc_number")
         doc_id = XdbFactory.getData(respQuery, "id")
         rev_id = XdbFactory.getData(respQuery, "rev_id")

         If (doc_code <> "") Then

            doc_name = Left(XdbFactory.getData(respQuery, "name") & " - " & XdbFactory.getData(respQuery, "description"), 80)

            issue = respQuery.fields.item("issue")
            issue = Xhelper.iff(IsNull(issue), "-1", issue)

            last_rev = respQuery.fields.item("last_rev")
            last_rev = Xhelper.iff(IsNull(last_rev), "-1", last_rev)





            doc_list.AddItem doc_id
            doc_list.List(doc_list.ListCount - 1, 1) = rev_id
            doc_list.List(doc_list.ListCount - 1, 2) = UCase(doc_code)
            doc_list.List(doc_list.ListCount - 1, 3) = last_rev
            doc_list.List(doc_list.ListCount - 1, 4) = issue
            doc_list.List(doc_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "status")
            doc_list.List(doc_list.ListCount - 1, 6) = doc_name
            doc_list.List(doc_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "category")
            doc_list.List(doc_list.ListCount - 1, 9) = file_name
            
            Call populate_who_doc_received_list(rev_id)
            
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

   header_titles = Array("ID", "Rev. ID", "Nº Documentos", "Rev", "TE", "Status", "Descrição")

   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, doc_list_header)



   Exit Function
error_handler:
   MsgBox "Erro: Documento Fora do Formato: " & varKey

End Function

Private Function move_files()

   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Dim file_name As String


   Load UserFormAlert
   UserFormAlert.Label1 = "Movendo os documentos"

   UserFormAlert.Show

   For i = 0 To doc_list.ListCount - 1

      doc_id = doc_list.List(i, 0)

      file_name = doc_list.List(i, 9)

      data("user_id") = auth.get_user_id
      data("review_id") = doc_list.List(i, 1)
      data("replace_date") = DateHelpers.FormatDateToSQlite(replace_date_txt.Value)

      Call db_documents_issued_replaced.Create(data)
      Call move_files_to_eng_folder(project_selected_id, doc_id, file_name)
   Next i

   Unload UserFormAlert
   Call Alert.Show("Documentos Modificados Com Sucesso", "", 2500)

   doc_list.Clear



End Function


Private Function move_files_to_eng_folder(ByVal project_selected_id As String, ByVal doc_id As String, file_name As String)

   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")

   Dim origin  As String
   origin = files_folder_path & "\" & file_name

   destiny = helper_folder_maker.get_eng_doc_folder(project_selected_id, doc_id, "SENT") & "\" & file_name

   If fso.FileExists(destiny) Then

      fso.DeleteFile destiny
      fso.moveFile origin, destiny

   End If



End Function







Private Function populate_who_doc_received_list(document_review_selected_id As String)



   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.get_who_received_the_documents(project_selected_id, document_review_selected_id)

   
   Do Until respQuery.EOF
   
   
   
   
    who_received_doc_list.AddItem XdbFactory.getData(respQuery, "doc_rev_id")
    
      who_received_doc_list.List(who_received_doc_list.ListCount - 1, 1) = "Enviar Mesma Revisão"
      who_received_doc_list.List(who_received_doc_list.ListCount - 1, 2) = UCase(XdbFactory.getData(respQuery, "received_company"))
      who_received_doc_list.List(who_received_doc_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "doc_number")
      who_received_doc_list.List(who_received_doc_list.ListCount - 1, 4) = "REV: [ " & XdbFactory.getData(respQuery, "rev") & " ] TE: [ " & XdbFactory.getData(respQuery, "te") & " ]  ST: [ " & XdbFactory.getData(respQuery, "status") & " ]"
      who_received_doc_list.List(who_received_doc_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "grd_number")
      who_received_doc_list.List(who_received_doc_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "grd_date")
      who_received_doc_list.List(who_received_doc_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "grd_description")
      
      respQuery.MoveNext
 

   Loop


   Dim header_titles As Variant

   header_titles = Array("ID", "Ação", "Enviado Para", "Nº Documentos", "REV x TE x STATUS", "GRD", "GRD Data", "GRD Nome")

   Call Xform.SetColumnWidthsAndHeader(who_received_doc_list, lblHidden, header_titles, who_received_doc_header_list)

End Function
