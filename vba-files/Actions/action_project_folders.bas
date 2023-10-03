Attribute VB_Name = "action_project_folders"

'namespace=vba-files\Actions


Public Function Create(ByVal project_id As String)


    Call create_project_folders_handler(project_id)

End Function


Public Function move_porject_files(ByVal project_id As String, ByVal folder_path As String)


    Call move_project_files(project_id, folder_path)

End Function



Private Function create_project_folders_handler(ByVal project_selected_id As String)


    Dim doc_id As String
    Dim contract_item As String
    Dim discipline_id As String
    Dim count_files As Long
    Dim count_files_folders As Long


    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")


    Dim respQuery As ADODB.Recordset
    Set respQuery = db_documents.get_all_project_documents(project_selected_id)

    Load UserFormAlert
    UserFormAlert.Label1.Caption = "Criando as pastas do Projeto (Emitidos)"
    UserFormAlert.Show

    count_files = 1
    count_files_folders = 1
    Do Until respQuery.EOF

        doc_id = XdbFactory.getData(respQuery, "id")
        contract_item = XdbFactory.getData(respQuery, "contract_item")
        discipline_id = XdbFactory.getData(respQuery, "discipline_id")

        If (doc_id <> "" And contract_item <> "" And discipline_id <> "") Then

            UserFormAlert.labelInfo.Caption = "Pasta: " & contract_item & "Arquivos:  [ " & count_files & " ] " & "  Pastas: [ " & count_files_folders & " ]"
            UserFormAlert.Repaint
            destiny = helper_folder_maker.get_eng_doc_folder(project_selected_id, doc_id, "SENT")

            count_files_folders = count_files_folders + 1
            Call Xhelper.waitMs(100)
        End If

        count_files = count_files + 1
        respQuery.MoveNext

    Loop

    UserFormAlert.Label1.Caption = "Pastas Criadas Com Sucesso"
    UserFormAlert.Repaint
    Call Xhelper.waitMs(2000)

    Unload UserFormAlert

End Function


Private Function move_project_files(ByVal project_selected_id As String, ByVal folder_path As String)


    Dim files_dict As Object
    Dim doc_id As String
    Dim doc_number As String
    Dim file_name_splited() As String
    Dim full_file_path_origin As String
    Dim count_files As Long
    Dim count_docs As Long
    count_docs = 0
    count_files = 0
    Set files_dict = CreateObject("Scripting.Dictionary")

      Set files_dict = file_helper.get_files_from_folders2(folder_path)


    Load UserFormAlert
    UserFormAlert.Label1.Caption = "Movendo os arquivos"
    UserFormAlert.Show

    For Each varKey In files_dict.Keys()
    
    If (varKey <> "count") Then
        file_search_name = files_dict(varKey)("file")
        file_extension = files_dict(varKey)("extension")
        file_name_splited = Split(UCase(file_search_name), "_REV_")
        doc_number = file_name_splited(0)
        full_file_path_origin = files_dict(varKey)("path")
          
        On Error GoTo ErrorHandler

      

            Dim respQuery As ADODB.Recordset
            Set respQuery = db_documents.SearchLimit(project_selected_id, doc_number, "doc_number")

            doc_code = XdbFactory.getData(respQuery, "doc_number")
            doc_id = XdbFactory.getData(respQuery, "id")
            contract_item = XdbFactory.getData(respQuery, "contract_item")
            discipline_id = XdbFactory.getData(respQuery, "discipline_id")
            count_docs = count_docs + 1
            
            If (doc_id <> "" And contract_item <> "" And discipline_id <> "") Then
            
                count_files = count_files + 1
                
                UserFormAlert.labelInfo.Caption = "Pasta: " & contract_item & "  Documentos:  [ " & count_docs & " ] " & "  Movidos: [ " & count_files & " ]"
                UserFormAlert.Repaint
                
                
                Call move_files_to_eng_folder(full_file_path_origin, project_selected_id, doc_id, UCase(file_search_name & "." & file_extension))
                
                Call Xhelper.waitMs(800)
                 UserFormAlert.Repaint
            End If

     End If

    Next '
    
    Unload UserFormAlert
    
ErrorHandler:

End Function


Private Function move_files_to_eng_folder(ByVal from As String, ByVal project_selected_id As String, ByVal doc_id As String, file_name As String)
   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")


   destiny = helper_folder_maker.get_eng_doc_folder(project_selected_id, doc_id, "SENT") & "\" & file_name

 
   Debug.Print from
   Debug.Print destiny
   
    If fso.FileExists(destiny) Then
     
  
      fso.DeleteFile destiny
   End If

   fso.moveFile from, destiny

End Function




Public Function move_project_file(ByVal project_selected_id As String, ByVal folder_path As String, selected_doc_numer As String)


 project_selected_id = 2
    Dim files_dict As Object
    Dim doc_id As String
    Dim doc_number As String
    Dim file_name_splited() As String
    Dim full_file_path_origin As String
    Dim count_files As Long
    Dim count_docs As Long
    count_docs = 0
    count_files = 0
    Set files_dict = CreateObject("Scripting.Dictionary")

      Set files_dict = file_helper.get_files_from_folders2(folder_path)


    Load UserFormAlert
    UserFormAlert.Label1.Caption = "Movendo os arquivos"
    UserFormAlert.Show

    For Each varKey In files_dict.Keys()
    
    If (varKey <> "count") Then
        file_search_name = files_dict(varKey)("file")
        file_extension = files_dict(varKey)("extension")
        file_name_splited = Split(UCase(file_search_name), "_REV_")
        doc_number = file_name_splited(0)
        full_file_path_origin = files_dict(varKey)("path")
          
        On Error GoTo ErrorHandler

      

            Dim respQuery As ADODB.Recordset
            Set respQuery = db_documents.SearchLimit(project_selected_id, doc_number, "doc_number")

            doc_code = XdbFactory.getData(respQuery, "doc_number")
            doc_id = XdbFactory.getData(respQuery, "id")
            contract_item = XdbFactory.getData(respQuery, "contract_item")
            discipline_id = XdbFactory.getData(respQuery, "discipline_id")
            count_docs = count_docs + 1
            
            If (doc_id <> "" And contract_item <> "" And discipline_id <> "" And selected_doc_numer = doc_code) Then
            
                count_files = count_files + 1
                
                UserFormAlert.labelInfo.Caption = "Pasta: " & contract_item & "  Documentos:  [ " & count_docs & " ] " & "  Movidos: [ " & count_files & " ]"
                UserFormAlert.Repaint
                
                
                Call move_files_to_eng_folder(full_file_path_origin, project_selected_id, doc_id, file_search_name & "." & file_extension)
                
                Call Xhelper.waitMs(800)
                 UserFormAlert.Repaint
            End If

     End If

    Next '
    
    Unload UserFormAlert
    
ErrorHandler:

End Function

