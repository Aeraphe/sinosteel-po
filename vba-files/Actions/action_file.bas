Attribute VB_Name = "action_file"

'namespace=vba-files\Actions

Public Function copy_to_temp_folder(file_path_origin As String, Optional foldername As String = "cdoc_tmp_files") As String


  
   
    Dim file_name As String
 
    Dim temp_file_path As String
    Dim integrity As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Temp_Fldr As String
    Temp_Fldr = fso.GetSpecialFolder(2)

    Dim cdoc_temp As String
    cdoc_temp = Temp_Fldr & "\" & foldername & "\"

    If Not fso.FolderExists(cdoc_temp) Then
        fso.CreateFolder cdoc_temp
    End If

    file_name = fso.GetFileName(file_path_origin)



    temp_file_path = cdoc_temp & "\" & file_name

    If fso.FileExists(temp_file_path) Then

        integrity = file_helper.checksum(file_path_origin, temp_file_path)

        If (Not integrity) Then
            Call fso.copyFile(file_path_origin, temp_file_path, True)
        End If
    Else
        Call fso.copyFile(file_path_origin, temp_file_path, True)
    End If

    copy_to_temp_folder = temp_file_path

End Function
