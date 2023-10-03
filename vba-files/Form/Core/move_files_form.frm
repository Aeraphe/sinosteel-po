VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} move_files_form 
   Caption         =   "Mover Arquivos (Organizar Pastas)"
   ClientHeight    =   13395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   24900
   OleObjectBlob   =   "move_files_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "move_files_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Core

Private project_selected_id As String
Private search_doc_option As String
Private doc_id_selected As String



Private Sub UserForm_Activate()

   Call auth.is_logged_to_access(Me)

End Sub

Private Sub btn_move_Click()
      Dim data As Object
      Set data = CreateObject("Scripting.Dictionary")
      Dim origin_path As String
      Dim doc_number As String
      Dim cat_id As String
      Dim i As Long
      
      Load UserFormAlert
      UserFormAlert.Label1.Caption = "Movendo os Arquivos"
      
    If (project_selected_id <> "") Then
        For i = 0 To move_files_list.ListCount - 1
        
           If move_files_list.Selected(i) = True Then
     
                origin_path = move_files_list.List(i, 4)
                doc_number = move_files_list.List(i, 1)
                cat_id = move_files_list.List(i, 3)
                data("category_id") = cat_id
                
                If (origin_path <> "" And doc_number <> "" And cat_id <> "") Then
                      UserFormAlert.labelInfo.Caption = "Documento: " & doc_number
                      UserFormAlert.Repaint
                      
                    Call update_handler(data, doc_id_selected)
                    Call move_project_file(project_selected_id, origin_path, doc_number)
                    
                    UserFormAlert.labelInfo.Caption = "Documento Movido: " & doc_number
                    UserFormAlert.Repaint
                    Call Xhelper.waitMs(400)
                
                End If
          End If
      Next i
    End If
End Sub

Private Function update_handler(data As Object, doc_id As String)

   Dim where As String

   If (doc_id <> "") Then

      where = "id='" & doc_id & "'"
      Call db_documents.update(data, where)

   Else
      MsgBox "Favor completar os daods", , "Dados incompletos"
   End If

End Function



Private Sub doc_list_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If (origin_path_txt <> "" And doc_number_txt.Value <> "" And doc_rev_txt.Value <> "" And doc_extension_txt.Value <> "") Then



Dim strFilePath As String

strFilePath = origin_path_txt.Value & "\" & doc_number_txt.Value & "_Rev_" & doc_rev_txt.Value & "." & LCase(doc_extension_txt.Value)

Call file_helper.open_file(strFilePath)

Else
  Call Alert.Show("Erro: Selecione um Documento antes", "", 2000)
End If
End Sub


Private Sub btn_open_folder_Click()

Dim folder_path As String
    folder_path = origin_path_txt.Value
 If (folder_path <> "") Then
     ' Shell "Explorer.exe " & MainFolder, vbNormalFocus
       ActiveWorkbook.FollowHyperlink Address:=folder_path, NewWindow:=True
     End If
  
End Sub

Private Sub btn_select_folder_Click()
Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .selectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then
      origin_path_txt.Value = sFolder
    End If

End Sub




Private Sub CommandButton1_Click()


If (doc_category_select.Value <> "") Then
  move_files_list.AddItem doc_list.List(doc_list.ListIndex, 0)
  move_files_list.List(move_files_list.ListCount - 1, 1) = doc_list.List(doc_list.ListIndex, 1)
  move_files_list.List(move_files_list.ListCount - 1, 2) = doc_list.List(doc_list.ListCount - 1, 4) & " ---->  " & doc_category_select.List(doc_category_select.ListIndex, 1)
  move_files_list.List(move_files_list.ListCount - 1, 3) = doc_category_select.List(doc_category_select.ListIndex, 0)
    move_files_list.List(move_files_list.ListCount - 1, 4) = origin_path_txt.Value
 
  
     Dim header_titles As Variant
   header_titles = Array("id", "Documento", "Categoria ->>", "id Categoria", "Origem")

   Call Xform.SetColumnWidthsAndHeader(move_files_list, lblHidden, header_titles, move_files_header_list)

   Call Xhelper.waitMs(400)
  End If
End Sub

Private Sub doc_list_Change()

Dim doc_id As String

   If (project_selected_id <> "") Then
      For i = 0 To doc_list.ListCount - 1

         If doc_list.Selected(i) = True Then

            doc_id_selected = doc_list.List(i, 0)
            If (doc_id_selected <> "") Then
                   origin_path_txt.Value = helper_folder_maker.get_eng_doc_folder(project_selected_id, doc_id_selected, "SENT")
                   
                   doc_number_txt.Value = doc_list.List(i, 1)
                   doc_rev_txt.Value = doc_list.List(i, 2)
                   doc_extension_txt.Value = doc_list.List(i, 8)
                   
                 
               Exit Sub
            End If
         End If
      Next i
   End If

End Sub





Private Sub origin_path_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .selectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then
      origin_path_txt.Value = sFolder
    End If

End Sub

Private Sub UserForm_Initialize()
 Call Shared_DocCategorySelectComp.Mount(doc_category_select)
  getDisciplineHandler
End Sub

Private Function getDisciplineHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = database_discipline.getAll()

   discipline_select.Clear
   Call Shared_CommonSelectComp.Mount(discipline_select, respQuery)
End Function


Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")


   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")


   End If
End Sub



Private Sub search_sec_title_btn_Click()
   search_doc_option = "description"
   Call SearchDocumentHandler(search_doc_option)
End Sub


Private Sub search_doc_btn_Click()
   search_doc_option = "name"
   Call SearchDocumentHandler(search_doc_option)
End Sub


Private Sub search_doc_sinosteel_btn_Click()
   search_doc_option = "sinosteel_doc_number"
   Call SearchDocumentHandler(search_doc_option)
End Sub


Private Sub search_doc_supplier_btn_Click()

   search_doc_option = "doc_number"
   Call SearchDocumentHandler(search_doc_option)

End Sub



Private Function SearchDocumentHandler(tb As String)

   Load UserFormAlert
   UserFormAlert.Label1.Caption = "Buscando Documentos"
   UserFormAlert.Show

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.SearchLastDocumentReview(project_selected_id, search_doc_txt.Value, tb)

   Dim doc_number As String

   doc_list.Clear

   Do Until respQuery.EOF

      doc_number = XdbFactory.getData(respQuery, "doc_number")

      doc_list.AddItem XdbFactory.getData(respQuery, "id")
      description = XdbFactory.getData(respQuery, "name") & " -- " & XdbFactory.getData(respQuery, "description")

      If (Len(description) <= 130) Then
         description_size = Len(description)
      Else
         description_size = title_size_txt.Value
      End If

      doc_list.List(doc_list.ListCount - 1, 1) = doc_number
      doc_list.List(doc_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "last_rev")
      doc_list.List(doc_list.ListCount - 1, 3) = Left(description, description_size)
      doc_list.List(doc_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "category")
      doc_list.List(doc_list.ListCount - 1, 5) = XdbFactory.getData(respQuery, "doc_type_code")
      doc_list.List(doc_list.ListCount - 1, 6) = XdbFactory.getData(respQuery, "discipline")
      doc_list.List(doc_list.ListCount - 1, 7) = XdbFactory.getData(respQuery, "contract_item")
      doc_list.List(doc_list.ListCount - 1, 8) = XdbFactory.getData(respQuery, "doc_extension")

      UserFormAlert.labelInfo.Caption = doc_number
      UserFormAlert.Repaint

      respQuery.MoveNext
   Loop

   UserFormAlert.Label1.Caption = "Busca Concluida"
   UserFormAlert.labelInfo.Caption = ""
   UserFormAlert.Repaint

   Call Xhelper.waitMs(2500)



   Dim header_titles As Variant
   header_titles = Array("id", "Documento", "Revisão", "Descrição", "Categoria", "Tipo", "Disciplina", "Contrato")

   Call Xform.SetColumnWidthsAndHeader(doc_list, lblHidden, header_titles, docs_header_listbox)

   Call Xhelper.waitMs(400)
   Unload UserFormAlert


End Function
