VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} adm_project_folders_form 
   Caption         =   "Administrar as pastas de arquivos do projeto"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18675
   OleObjectBlob   =   "adm_project_folders_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "adm_project_folders_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Adm_Files


Private project_selected_id


Private Sub UserForm_Activate()

    Call auth.is_logged_to_access(Me)
    Call auth.is_authorized("SUPER_ADMIN", Me)
 
 End Sub


Private Sub btn_move_files_Click()

    adm_project_folders_form.Hide

    Call action_project_folders.move_porject_files(project_selected_id, files_path_txt.Value)
    adm_project_folders_form.Show
End Sub

Private Sub btn_move_project_files_Click()
    Dim folder_path As String
    folder_path = file_helper.open_folder_dialog
    If (folder_path <> "" And project_selected_id <> "") Then
       Call action_organize_project_files.start(project_selected_id, folder_path)
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
        Call validate_frames(True)

        Else
           Call validate_frames(False)
    End If
    
End Sub

Private Function validate_frames(ByVal state As Boolean)
  actions_fr.Enabled = state
  compare_files_fr.Enabled = state
  move_files_fr.Enabled = state
End Function

Private Sub btn_create_sent_folders_Click()
    adm_project_folders_form.Hide

    Call action_project_folders.Create(project_selected_id)
    adm_project_folders_form.Show
End Sub

Private Sub origin_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    origin_txt.Value = open_dialog_select_folder_handler

End Sub

Private Sub destiny_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    destiny_txt.Value = open_dialog_select_folder_handler
End Sub

Private Function open_dialog_select_folder_handler() As String
    Dim sFolder As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            sFolder = .selectedItems(1)
        End If
    End With

    open_dialog_select_folder_handler = sFolder

End Function


Private Sub btn_compar_Click()

    ScreenUpdating = Fasle
    Dim originFiles As Object
    Dim destinyFiles As Object

    Set originFiles = file_helper.get_files_from_folders2(origin_txt.Value)
    Set destinyFiles = file_helper.get_files_from_folders2(destiny_txt.Value)
    Call compare_files(originFiles, destinyFiles)
    ScreenUpdating = True
End Sub


Private Function create_report(originFiles As Object, destinyFiles As Object, nextRow As Long)
    Dim tableObject As ListObject

    total_rows = receive_x_sent_sheet.Range("RECEIVED_X_SENT_TB").Rows.count
    Set tableObject = receive_x_sent_sheet.ListObjects("RECEIVED_X_SENT_TB")

End Function

Private Function compare_files(originFiles As Object, destinyFiles As Object)

    Dim file_name_origin As String
    Dim file_name_destiny As String
    Dim file_path_origin As String
    Dim file_path_destiny As String
    Dim file_full_path_origin As String
    Dim file_full_path_destiny As String

    Dim tableObject As ListObject
    Dim nextRow As Long
    Dim doc_finded As Boolean
    Dim varKey2 As Variant
    Dim varKey As Variant

    nextRow = 1
    i = 1
    j = 1
    Set tableObject = receive_x_sent_sheet.ListObjects("RECEIVED_X_SENT_TB")
    tableObject.DataBodyRange.ClearContents

    For Each varKey In originFiles.Keys()

        doc_finded = False

        If (varKey <> "count") Then
            file_name_origin = UCase(originFiles(varKey)("file"))
            file_path_origin = originFiles(varKey)("folder")
            file_full_path_origin = originFiles(varKey)("path")



            Set originFile = is_file_in_correct_format(file_name_origin)

            If (originFile("STATUS")) Then

                file_origin_name_lb.Caption = UCase(originFile("DOC"))

                For Each varKey2 In destinyFiles.Keys()
                    If (varKey2 <> "count" And varKey2 <> "") Then

                        file_name_destiny = UCase(destinyFiles(varKey2)("file"))
                        file_path_destiny = destinyFiles(varKey2)("folder")
                        file_full_path_destiny = destinyFiles(varKey2)("path")

                        Set destinyFile = is_file_in_correct_format(file_name_destiny)
                        If (destinyFile("STATUS")) Then
                            file_destiny_name_lb.Caption = UCase(destinyFile("DOC"))
                            Me.Repaint

                            If (file_name_origin = file_name_destiny) Then


                                tableObject.ListColumns("ARQUIVO").DataBodyRange(nextRow).Value = file_name_origin
                                tableObject.ListColumns("LOCAL DO ARQUIVO RECEBIDOS").DataBodyRange(nextRow).Value = file_path_origin
                                tableObject.ListColumns("LOCAL DO ARQUIVO EMITIDOS").DataBodyRange(nextRow).Value = file_path_destiny
                                tableObject.ListColumns("STATUS").DataBodyRange(nextRow).Value = "LOCALIZADO"
                                If (check_signature_ck.Value) Then
                                    file_origin_md5 = file_helper.FileToMD5Hex(file_full_path_origin)
                                    file_destiny_md5 = file_helper.FileToMD5Hex(file_full_path_destiny)
                                    tableObject.ListColumns("MESMO ARQUIVO?").DataBodyRange(nextRow).Value = Xhelper.iff(file_origin_md5 = file_destiny_md5, "MESMO ARQUIVO", "ARQUIVO DIFERENTE")
                                Else
                                    tableObject.ListColumns("MESMO ARQUIVO?").DataBodyRange(nextRow).Value = "VERIFICAÇÂO DE ASSINATURA DESABILITADA"
                                End If
                                doc_finded = True
                                nextRow = nextRow + 1
                            End If
                        End If


                    End If
                Next varKey2
            End If

            If (Not doc_finded) Then
                tableObject.ListColumns("ARQUIVO").DataBodyRange(nextRow).Value = file_name_origin
                tableObject.ListColumns("LOCAL DO ARQUIVO RECEBIDOS").DataBodyRange(nextRow).Value = file_path_origin
                tableObject.ListColumns("LOCAL DO ARQUIVO EMITIDOS").DataBodyRange(nextRow).Value = ""
                tableObject.ListColumns("STATUS").DataBodyRange(nextRow).Value = "NÂO LOCALIZADO"
                nextRow = nextRow + 1
            End If
        End If



    Next varKey

End Function


Private Function is_file_in_correct_format(file_name As String) As Object

    Set response = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrorHanlder
    file_name_splited = Split(UCase(file_name), "_REV_")
    response("DOC") = file_name_splited(0)
    response("STATUS") = True

    Set is_file_in_correct_format = response
    Exit Function

ErrorHanlder:
    response("DOC") = "ERROR"
    response("STATUS") = False
    Set is_file_in_correct_format = response
End Function
