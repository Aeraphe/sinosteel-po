VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_import_form 
   Caption         =   "Importar Lista de Documentos"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   OleObjectBlob   =   "doc_import_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_import_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents

Public project_selected_id
Private import_app As Excel.Application
Private import_book As Excel.Workbook
Private import_sheet As Worksheet
Private import_file_info As Object


Private Sub UserForm_Activate()


   Call auth.is_logged_to_access(Me)
   Call auth.is_authorized("SUPER_ADMIN", Me)

End Sub

Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")
      fr_get_file.Enabled = True

   Else

      fr_get_file.Enabled = False
   End If

End Sub

Private Sub btn_get_file_Click()

   Dim strFile As String
   Set import_file_info = CreateObject("Scripting.Dictionary")

   strFile = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsb*), *.xlsx*", title:="Escolha um Arquivo", MultiSelect:=False)

   If (strFile <> "") Then

      import_file_info("FULL_FILE_PATH") = strFile
      With CreateObject("Scripting.FileSystemObject")
         import_file_info("fileName") = .GetFileName(strFile)
         import_file_info("extName") = .GetExtensionName(strFile)
         import_file_info("baseName") = .GetBaseName(strFile)
         import_file_info("parentName") = .GetParentFolderName(strFile)
      End With
      lb_file_selected.Caption = import_file_info("fileName")

      Call enable_frames(True)
   Else
      Call enable_frames(False)
   End If

End Sub

Private Function close_hidden_excel(Optional app_to_close As String = "Excel.exe")

   Dim closeApp As String

   Set app = CreateObject("Excel.Sheet")
   app.Application.Visible = True

   closeApp = "TASKKILL /F /IM " & app_to_close
   Shell closeApp, vbHide
   MsgBox "Excel object closed"

End Function

Private Function enable_frames(Optional state As Boolean = False)
   fr_get_file.Enabled = state
   fr_actions.Enabled = state
   fr_change_doc_info.Enabled = state
   fr_update_doc_prop.Enabled = state
End Function




Private Function load_import_excel_app(full_file_path As String)

   Set import_app = New Excel.Application
   import_app.Visible = False 'Visible is False by default, so this isn't necessary
   Set import_book = import_app.Workbooks.Add(full_file_path)
   Set import_sheet = import_book.Sheets("index")

End Function

Private Sub import_doc_btn_Click()
   Application.ScreenUpdating = False
   
   Call load_import_excel_app(import_file_info("FULL_FILE_PATH"))
   import_documents_handler
   
   Application.DisplayAlerts = False
   import_book.Save
   import_book.Close
   import_app.Quit
   Application.DisplayAlerts = True
   
   Application.ScreenUpdating = True
End Sub




Private Function import_documents_handler()


   Dim total_rows As Long
   Dim i  As Long
   Dim prop_name As String
   Dim prop_value  As Variant
   Dim doc_id As String
   Dim iObject As ListObject
   Dim iNewRow As ListRow
   Dim data As Object
   Dim total_inport As Long
   Dim doc_number As String
   Dim discipline_id As String
   Dim category_id As String
   

   Load UserFormAlert
   UserFormAlert.Label1.Caption = "Importando Documentos"
   UserFormAlert.Show


   total_rows = import_sheet.Range("import_documents_table").Rows.count
   Set iObject = import_sheet.ListObjects("import_documents_table")


 
   Set data = CreateObject("Scripting.Dictionary")

   data("project_id") = project_selected_id

   total_inport = 1

   For i = 1 To total_rows

      If (iObject.ListColumns("Numero_Fornecedor").DataBodyRange(i).Value <> "") Then

         doc_number = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Numero_Fornecedor").DataBodyRange(i).Value))
         
         discipline_id = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("disciplina_id").DataBodyRange(i).Value))
         If (discipline_id <> "") Then
          data("discipline_id") = discipline_id
         End If
         
        category_id = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("categoria_id").DataBodyRange(i).Value))
       
         If (category_id <> "") Then
          data("category_id") = category_id
         End If
         

         UserFormAlert.Label1.Caption = "Importando Documentos : " & total_inport & " : " & total_rows - i
         UserFormAlert.labelInfo.Caption = doc_number
         UserFormAlert.Repaint


         data("doc_number") = doc_number
         data("sinosteel_doc_number") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Numero_Sinosteel").DataBodyRange(i).Value))
         data("name") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Titulo_Primario").DataBodyRange(i).Value))
         data("description") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Titulo_Secundario").DataBodyRange(i).Value))
         data("pages") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Paginas").DataBodyRange(i).Value))
         data("doc_type_code") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Codigo_Documento").DataBodyRange(i).Value))
         data("doc_format") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Formato").DataBodyRange(i).Value))
         data("contract_item") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Item_Contrato").DataBodyRange(i).Value))
         data("doc_extension") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Extensao").DataBodyRange(i).Value))

         total_inport = total_inport + 1
         doc_id = db_documents.Import(data)

         If (doc_id <> "") Then

            UserFormAlert.Label1.Caption = "Incluindo Propriedade : " & total_inport & " : " & total_rows - i
            UserFormAlert.Repaint

            iObject.ListColumns("ID").DataBodyRange(i) = doc_id
            prop_name = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Propriedade").DataBodyRange(i).Value))
            prop_value = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Valor").DataBodyRange(i).Value))
            Call insert_property_handler(doc_id, prop_name, prop_value)
         End If

         UserFormAlert.Label1.Caption = "Incluindo Revisão : " & total_inport & " : " & total_rows - i
         UserFormAlert.Repaint
         Call insert_doc_review_handler(doc_id, iObject, i)
      End If

   Next i

End Function




Private Sub btn_import_reviews_Click()

   Call load_import_excel_app(import_file_info("FULL_FILE_PATH"))

   Dim tbObject As ListObject
   Dim iNewRow As ListRow
   Dim doc_id As String
   Dim doc_rev As String
   Dim doc_issue As String
   Dim i As Long
   Dim answer As Integer
   Dim total_inport_rev As Long
   total_inport_rev = 0
   answer = MsgBox("Gostaria de importar as Revisões?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")


   If (answer = vbYes And project_selected_id <> "") Then
      Load UserFormAlert
      UserFormAlert.Label1.Caption = "Importando Revisões"
      UserFormAlert.Show

      Set tbObject = import_sheet.ListObjects("import_documents_table")

      total_rows = import_sheet.Range("import_documents_table").Rows.count

      For i = 1 To total_rows
         doc_id = tbObject.ListColumns("ID").DataBodyRange(i).Value
         doc_rev = helper_string.RemoveLineBreak(UCase(tbObject.ListColumns("Revisao").DataBodyRange(i).Value))
         doc_issue = helper_string.RemoveLineBreak(UCase(tbObject.ListColumns("Emissao").DataBodyRange(i).Value))
         doc_number = helper_string.RemoveLineBreak(UCase(tbObject.ListColumns("Numero_Fornecedor").DataBodyRange(i).Value))



         If (Not check_doc_review_exist(doc_id, doc_rev, doc_issue)) Then
            Call insert_doc_review_handler(doc_id, tbObject, i)
            total_inport_rev = total_inport_rev + 1
         End If

         UserFormAlert.Label1.Caption = "Importando Revisões: " & total_inport_rev & " : " & total_rows - i
         UserFormAlert.labelInfo.Caption = doc_number & " REV: " & doc_rev & " Emissão: " & doc_issue
         UserFormAlert.Repaint

      Next i
   End If
   UserFormAlert.Label1.Caption = "Finalizado"
   UserFormAlert.labelInfo.Caption = "Total de Revisões importadas: " & total_inport_rev

   import_app.Quit

End Sub


Private Function check_doc_review_exist(doc_id As String, review_code As String, issue As String) As Boolean

   If (doc_id <> "" And review_code <> "" And issue <> "") Then

      Dim respQuery As ADODB.Recordset
      Set respQuery = db_documents.get_doc_review_issue(doc_id, review_code, issue)

      If (XdbFactory.getData(respQuery, "doc_id")) Then
         check_doc_review_exist = True
      Else
         check_doc_review_exist = False
      End If

   Else
      check_doc_review_exist = True

   End If

End Function

Private Function insert_doc_review_handler(doc_id As String, iObject As ListObject, row_number As Long)


   Dim doc_rev As String
   Dim doc_issue As String


   If (doc_id <> "") Then

      grd_data = ""
      grd_status_date = ""
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      On Error Resume Next
      grd_data = format(CDate(Trim(UCase(iObject.ListColumns("Grd_Data").DataBodyRange(row_number).Value))), "YYYY-MM-DD")

      On Error Resume Next
      grd_status_date = format(CDate(Trim(UCase(iObject.ListColumns("Status_Grd_Data").DataBodyRange(row_number).Value))), "YYYY-MM-DD")



      doc_rev = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Revisao").DataBodyRange(row_number).Value))
      doc_issue = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Emissao").DataBodyRange(row_number).Value))
      doc_number = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Numero_Fornecedor").DataBodyRange(row_number).Value))
      doc_status = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Status").DataBodyRange(row_number).Value))

      If (doc_rev <> "" And doc_issue <> "") Then



         data("user_id") = auth.get_user_id
         data("doc_id") = doc_id
         data("rev_code") = doc_rev
         data("issue") = doc_issue
         data("grd") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Rev_Grd").DataBodyRange(row_number).Value))
         data("grd_date") = grd_data
         data("status") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Status").DataBodyRange(row_number).Value))
         data("grd_status") = doc_status
         data("grd_status_date") = grd_status_date
         data("file_name") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Arquivo").DataBodyRange(row_number).Value))
         data("file_extension") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Extensao").DataBodyRange(row_number).Value))
         data("obs") = helper_string.RemoveLineBreak(UCase(iObject.ListColumns("Obs").DataBodyRange(row_number).Value))





         Call db_documents.InsertDocumentReview(data)


      End If


   End If



End Function




Private Function insert_property_handler(doc_id As String, prop_name As String, prop_value As Variant)

   If (prop_name <> "" And doc_id <> "") Then


      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")

      Dim respQuery As ADODB.Recordset
      Set respQuery = db_document_props.SearchType(prop_name)
      prop_id = respQuery.fields("id").Value

      If (prop_id <> "") Then

         data("document_id") = doc_id
         data("property_id") = prop_id
         data("value") = prop_value

         Call db_document_props.Create(data)
      End If

   End If
End Function
