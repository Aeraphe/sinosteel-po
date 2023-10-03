Attribute VB_Name = "shared_project_files"



'namespace=vba-files\Shared


Function copyFile(ByVal SourceFilePath As String, ByVal project_selected_id As String, ByVal doc_id As String, fileName As String) As Boolean

    ' Declare variables to be used in the function
    Dim fileWasCopied As Boolean
    Dim msgResponse As String
    Dim destiny As String

    ' Use helper_folder_maker to get the SENT folder for the specified document in the specified project
    destiny = helper_folder_maker.get_eng_doc_folder(project_selected_id, doc_id, "SENT")

    ' Call the CopyFileLongPathWithChecksum function to copy the file from the SourceFilePath to the destiny location
     fileWasCopied = h_windows_api_file.CopyFileLongPathWithChecksum(SourceFilePath, destiny, False)

    ' Check if the file was successfully copyed and create status message accordingly
    If (fileWasCopied) Then
        msgResponse = "Copy File (" & fileName & ") Successfully!"
    Else
        msgResponse = "Copy File (" & fileName & ") Error!"
    End If

    ' Call the helper_log function to write the response status to the debug log
    Call helper_log.createLogFile(msgResponse, DEGUB_COPY_FILE_NAME)

    copyFile = fileWasCopied
End Function





'/*
' This function moves a file from a source folder to a project's specified folder (defaulting to SENT)
' It also logs any errors or successes in the application's debug log
'*/
Function moveFile(ByVal SourceFilePath As String, ByVal project_selected_id As String, ByVal doc_id As String, ByVal fileName As String, Optional ByVal folderType As String = "SENT")

    ' Declare variables to be used in the function
    Dim fileWasCopied As Boolean
    Dim msgResponse As String
    Dim folderDestiny As String

    ' Use helper_folder_maker to get the specified folder for the document in the project
    folderDestiny = helper_folder_maker.get_eng_doc_folder(project_selected_id, doc_id, folderType)

    fileWasCopied = h_windows_api_file.CopyFileLongPathWithChecksum(SourceFilePath, folderDestiny, True)

    ' Check if the file was successfully moved and create status message accordingly
    If (fileWasCopied) Then
        msgResponse = "Move File (" & fileName & ") Successfully!"
    Else
        msgResponse = "Move File (" & fileName & ") Error!"
    End If

    ' Call the helper_log function to write the response status to the debug log
    Call helper_log.DebugApp(msgResponse)

End Function






Private Function prepar_files_detiny(ByVal SourceFilePath As String) As Object

    Dim fso As Object
    Dim Temp_Fldr As String
    Dim file_name As String
    Dim temp_moved_project_folder As String
    Dim temp_copied_folder_full_path As String
    Dim response As Object

    Set response = CreateObject("Scripting.Dictionary")

    Set fso = CreateObject("Scripting.FileSystemObject")

    Temp_Fldr = fso.GetSpecialFolder(2)
    file_name = fso.GetFileName(SourceFilePath)
    temp_moved_project_folder = "cdoc_moved_file" & format(Now(), "_yyyy_MM_dd")
    temp_copyed_project_folder = "cdoc_copied_file" & format(Now(), "_yyyy_MM_dd")

    temp_moved_folder_full_path = Temp_Fldr & "\" & temp_moved_project_folder & "\"
    temp_copied_folder_full_path = Temp_Fldr & "\" & temp_copyed_project_folder & "\"

    Debug.Print temp_moved_project_folder
    Debug.Print temp_copyed_project_folder


    Debug.Print temp_moved_folder_full_path
    Debug.Print temp_copied_folder_full_path

    If Not fso.FolderExists(temp_moved_folder_full_path) Then


        fso.CreateFolder temp_moved_folder_full_path

    End If

    If Not fso.FolderExists(temp_copied_folder_full_path) Then

        fso.CreateFolder temp_copied_folder_full_path

    End If

    If fso.FileExists(temp_moved_folder_full_path & "\" & file_name) Then
        fso.DeleteFile temp_moved_folder_full_path & "\" & file_name
    End If

    If fso.FileExists(temp_copied_folder_full_path & "\" & file_name) Then
        fso.DeleteFile temp_copied_folder_full_path & "\" & file_name
    End If

    response("MOVE_PATH") = temp_moved_folder_full_path & "\" & file_name
    response("COPY_PATH") = temp_copied_folder_full_path & "\" & file_name



    Set prepar_files_detiny = response

End Function
