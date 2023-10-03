Attribute VB_Name = "file_helper"



'namespace=vba-files\Helpers


'/*
'
'
'Helper for copy a sheet to new Sheet File
'
'*/
Public Function copy_selected_sheet_to_new_file(copieWorkSheet As Object)
    copieWorkSheet.Activate
    copieWorkSheet.Cells.Select
    Selection.Copy

    Dim LTEGenerateWB As Workbook
    Set LTEGenerateWB = Application.ActiveWorkbook

    Dim wb As Workbook
    Set wb = Workbooks.Add

    Windows(wb.name).Activate
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    ActiveSheet.Paste

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False



End Function






Private Function SaveLTE()

    Dim colorLTE As Variant

    colorLTE = Split(LTESheet.Range("LTE_FIRST_VOL").Value, "-")
    Dim LTEFilename As String

    LTEFilename = ConfigSheet.Range("CONFIG_LTE_FIRST_NAME").Value & "_" & LTESheet.Range("MAT_ORIGIN").Value & "_" & Replace(LTESheet.Range("LTE_N").Value, "/", "_") & "_" & colorLTE(2)
    Dim fileFolder As String
    fileFolder = ConfigSheet.Range("CONFIG_LTE_FILE_PATH").Value

    Dim file_full_path  As String

    file_full_path = fileFolder & "\" & LTEFilename & ".xlsx"

    'Delete File if exist
    Call DeleteFile(file_full_path)

    ActiveWorkbook.SaveAs fileName:=file_full_path, FileFormat:= _
    xlOpenXMLWorkbook, CreateBackup:=False

    ActiveWorkbook.Close SaveChanges:=False
End Function

Function FileExists(ByVal FileToTest As String) As Boolean
    FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
    If FileExists(FileToDelete) Then 'See above
        ' First remove readonly attribute, if set
        SetAttr FileToDelete, vbNormal
        ' Then delete the file
        Kill FileToDelete
    End If
End Sub



Function get_files_from_folders(folderPath As String) As Object

    Dim fso As Object
    Dim FSfolder As Object, FSsubfolder As Object, FSfile As Object
    Dim folders As Collection, levels As Collection
    Dim subfoldersColl As Collection
    Dim n As Long, c As Long, i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folders = New Collection
    Set levels = New Collection

    'Add start folder to stack

    folders.Add fso.GetFolder(folderPath)
    levels.Add 0

    n = 0

    Dim response As Object
    Set response = CreateObject("Scripting.Dictionary")

    Do While folders.count > 0

        'Remove next folder from top of stack

        Set FSfolder = folders(folders.count): folders.Remove folders.count
        c = levels(levels.count): levels.Remove levels.count

        'Output this folder and its files


        n = n + 1
        c = c + 1
        For Each FSfile In FSfolder.Files
            On Error Resume Next
            If (FSfile.Attributes And 2) <> 2 Then
                response(FSfile.name) = FSfile.name
                n = n + 1
            End If

        Next FSfile
        response("count") = n - 1
        'Get collection of subfolders in this folder

        Set subfoldersColl = New Collection
        For Each FSsubfolder In FSfolder.SubFolders
            subfoldersColl.Add FSsubfolder
        Next FSsubfolder

        'Loop through collection in reverse order and put each subfolder on top of stack.  As a result, the subfolders are processed and
        'output in the correct ascending ASCII order

        For i = subfoldersColl.count To 1 Step -1
            If folders.count = 0 Then
                folders.Add subfoldersColl(i)
                levels.Add c
            Else
                folders.Add subfoldersColl(i), , , folders.count
                levels.Add c, , , levels.count
            End If
        Next i
        Set subfoldersColl = Nothing

        Loop 'next

        Set get_files_from_folders = response

End Function



Function get_files_from_folders2(folderPath As String) As Object

    Dim fso As Object
    Dim FSfolder As Object, FSsubfolder As Object, FSfile As Object
    Dim folders As Collection, levels As Collection
    Dim subfoldersColl As Collection
    Dim n As Long, c As Long, i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folders = New Collection
    Set levels = New Collection

    'Add start folder to stack

    folders.Add fso.GetFolder(folderPath)
    levels.Add 0

    n = 0

    Dim response As Object
    Set response = CreateObject("Scripting.Dictionary")

    Do While folders.count > 0

        'Remove next folder from top of stack

        Set FSfolder = folders(folders.count): folders.Remove folders.count
        c = levels(levels.count): levels.Remove levels.count

        'Output this folder and its files


        n = n + 1
        c = c + 1
        For Each FSfile In FSfolder.Files
            On Error Resume Next


            Dim file_info As Object
            Set file_info = CreateObject("Scripting.Dictionary")
            If (FSfile.Attributes And 2) <> 2 Then
                file_info("file") = fso.GetBaseName(FSfile.name)
                file_info("extension") = fso.GetExtensionName(FSfile.path)
                file_info("path") = FSfile.path
                file_info("folder") = fso.GetParentFolderName(FSfile.path)
                item = "i" & n
                response.Add item, file_info
                ss = response(item)("file")
                n = n + 1
            End If
        Next FSfile
        response("count") = n - 1
        'Get collection of subfolders in this folder

        Set subfoldersColl = New Collection
        For Each FSsubfolder In FSfolder.SubFolders
            subfoldersColl.Add FSsubfolder
            Next

            'Loop through collection in reverse order and put each subfolder on top of stack.  As a result, the subfolders are processed and
            'output in the correct ascending ASCII order

            For i = subfoldersColl.count To 1 Step -1
                If folders.count = 0 Then
                    folders.Add subfoldersColl(i)
                    levels.Add c
                Else
                    folders.Add subfoldersColl(i), , , folders.count
                    levels.Add c, , , levels.count
                End If
                Next
                Set subfoldersColl = Nothing

            Loop

            Set get_files_from_folders2 = response

End Function

Function move_files_from_folder(FromPath As String, ToPath As String)
    Dim fso As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim FileInFromFolder As Object


    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each FileInFromFolder In fso.GetFolder(FromPath).Files
        FileInFromFolder.Move ToPath
    Next FileInFromFolder

End Function


Public Function open_file(ByVal strFilePath As String)

    On Error GoTo ErrorHandler
    If (strFilePath <> "") Then
        Load UserFormAlert
        UserFormAlert.Label1 = "Abrindo Documento!!!"
        UserFormAlert.Show
        Call ActiveWorkbook.FollowHyperlink(strFilePath)

        Unload UserFormAlert
    End If
ErrorHandler:

End Function

Public Function open_folder(ByVal folder_path As String)
    On Error GoTo ErrorHandler
    If (folder_path <> "") Then
        ' Shell "Explorer.exe " & MainFolder, vbNormalFocus
        Load UserFormAlert
        UserFormAlert.Label1 = "Abrindo Pasta Documento!!!"
        UserFormAlert.Show
        ActiveWorkbook.FollowHyperlink Address:=folder_path, NewWindow:=True
        Unload UserFormAlert
    End If
ErrorHandler:

End Function


Public Function open_folder_dialog() As String

    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .selectedItems(1)
        End If
    End With

    open_folder_dialog = sFolder
End Function



Public Function open_file_dialog() As String

    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogOpen)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .selectedItems(1)
        End If
    End With

    open_file_dialog = sFolder
End Function




Public Function FileToMD5Hex(sFileName As String) As String
    Dim enc
    Dim bytes
    Dim outstr As String
    Dim pos As Integer

    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFileName)
    bytes = enc.ComputeHash_2((bytes))
    'Convert the byte array to a hex string
    For pos = 1 To LenB(bytes)
        outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
        Next
        FileToMD5Hex = outstr
        Set enc = Nothing
End Function

Public Function FileToSHA1Hex(sFileName As String) As String
    Dim enc
    Dim bytes
    Dim outstr As String
    Dim pos As Integer

    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFileName)
    bytes = enc.ComputeHash_2((bytes))
    'Convert the byte array to a hex string
    For pos = 1 To LenB(bytes)
        outstr = outstr & LCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
        Next
        FileToSHA1Hex = outstr 'Returns a 40 byte/character hex string
        Set enc = Nothing
End Function


Private Function GetFileBytes(ByVal path As String) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    lngFileNum = FreeFile


    If LenB(Dir(path)) Then ''// Does file exist?
        Open path For Binary Access Read As lngFileNum
        ReDim bytRtnVal(LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function


'/*
'
'
'Check the integrity of file by MD5 Hash
'
'
'*/
Public Function checksum(ByVal source_path As String, ByVal destiny_path As String) As Boolean

    Dim source_hex As String
    Dim destiny_hex As String


    source_hex = FileToMD5Hex(source_path)
    destiny_hex = FileToMD5Hex(destiny_path)



    If (source_hex = destiny_hex) Then

        checksum = True

    Else
        checksum = False
    End If

End Function




Function IsFileOpen(fileName As String) As Boolean
    Dim fileNum As Integer
    Dim errNum As Integer

    'Allow all errors to happen
    On Error Resume Next
    fileNum = FreeFile()

    'Try to open and close the file for input.
    'Errors mean the file is already open
    Open fileName For Input Lock Read As #fileNum
    Close fileNum

    'Get the error number
    errNum = Err

    'Do not allow errors to happen
    On Error GoTo 0

    'Check the Error Number
    Select Case errNum

        'errNum = 0 means no errors, therefore file closed
     Case 0
        IsFileOpen = False

        'errNum = 70 means the file is already open
     Case 70
        IsFileOpen = True

        'Something else went wrong
     Case Else
        IsFileOpen = errNum

    End Select

End Function





Public Function moveFilesWithCheckSum(ByVal origin As String, ByVal destiny As String, Optional Try As Long = 1) As Boolean


    On Error GoTo ErrorHandler

    Dim success As Boolean
    Dim fso As Object
    Dim debugInfo As String


    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if the source file exists
    If Not fso.FileExists(origin) Then
        MsgBox "Arquivo não Existe"
        debugInfo = "Arquivo de Origem Não Existe: " & origin
        moveFilesWithCheckSum = False
        Call helper_log.DebugApp(debugInfo)
        Exit Function

    Else
        debugInfo = "Arquivo de Origem Existe: " & origin
    End If

    ' Create destination folder if it doesn't exist
    If (CreateFolder(destiny)) Then
        debugInfo = debugInfo & vbNewLine & " Pasta de destino Criada com Sucesso"
    End If

    success = False
    Do While Not success And Try <= 5
        ' Copy file and check for checksum integrity
        Call fso.copyFile(origin, destiny, True)
        success = file_helper.checksum(origin, destiny)
        If (success) Then
            debugInfo = debugInfo & vbNewLine & "Movido Com Sucesso: " & origin
        Else
            debugInfo = debugInfo & vbNewLine & "Não Foi Movido: " & origin
        End If
        ' Retry if unsuccessful
        Try = Try + 1
    Loop

    ' Delete original file if checksum integrity was achieved
    If success Then
        Call fso.DeleteFile(origin)
    End If

    ' Return result
    moveFilesWithCheckSum = success
    Call helper_log.DebugApp(vbNewLine & debugInfo)
    Exit Function

ErrorHandler:

    Dim errorMsg As String
    errorMsg = "Ocorreu um erro ao Mover os Arquivos " & Err.description
    MsgBox errorMsg
    Call helper_log.DebugApp(vbNewLine & errorMsg & vbNewLine & debugInfo)
End Function



Public Function copyFilesWithCheckSum(ByVal origin As String, ByVal destiny As String, Optional ByVal maxTries As Long = 5) As Boolean
    On Error GoTo ErrorHandler

    ' Declare variables with appropriate data types
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim integrity As Boolean: integrity = False
    Dim tryCount As Long: tryCount = 1
    Dim debugInfo As String

    ' Check if the source file exists
    If Not fso.FileExists(origin) Then

        debugInfo = "Arquivo de Origem Não Existe: " & origin
        MsgBox "Arquivo de Origem Não Existe"
        copyFilesWithCheckSum = False
        Call helper_log.DebugApp(vbNewLine & debugInfo)
        Exit Function

    Else
        debugInfo = "Arquivo de Origem Existe"
    End If

    ' Create destination folder if it doesn't exist
    If (CreateFolder(destiny)) Then
        debugInfo = debugInfo & vbNewLine & " Pasta de destino Criada com Sucesso"
    End If


    ' Copy the file and check its checksum
    Do While Not integrity And tryCount <= maxTries
        fso.copyFile origin, destiny, True
        integrity = file_helper.checksum(origin, destiny)
        debugInfo = debugInfo & vbNewLine & "[" & tryCount & "] Arquivo Copiado: " & origin
        tryCount = tryCount + 1
    Loop

    ' Return result
    copyFilesWithCheckSum = integrity
    Call helper_log.DebugApp(vbNewLine & debugInfo)

    Exit Function

ErrorHandler:

    Dim errorMsg As String
    errorMsg = "Ocorreu um erro ao Copiar os Arquivos " & Err.description
    MsgBox errorMsg
    Call helper_log.DebugApp(vbNewLine & errorMsg & vbNewLine & debugInfo)
End Function



Function CreateFolder(path As String) As Boolean
    Dim fso As Object
    Dim arrFolders() As String
    Dim folderPath As String
    Dim i As Integer
    Dim debugInfo As String

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Split the path into individual folder names
    arrFolders = Split(path, "\")

    ' Initialize the folder path to the first folder name
    folderPath = arrFolders(0) & "\"

    ' Loop through each folder name and create the folder if it doesn't exist
    For i = 1 To UBound(arrFolders)
        ' If we have reached the last folder name and it contains a file extension,
        ' exit the loop since we don't need to create a folder for this file
        If InStr(arrFolders(i), ".") > 0 And i = UBound(arrFolders) Then
            Exit For
        Else
            folderPath = folderPath & arrFolders(i) & "\"

            ' Create the folder if it doesn't exist
            If Not fso.FolderExists(folderPath) Then
                fso.CreateFolder folderPath
            End If
        End If
    Next i

    CreateFolder = True

    Exit Function

ErrorHandler:
    debugInfo = "Error on creating folder: " & Err.description
    CreateFolder = False
    Call helper_log.DebugApp(debugInfo)
End Function

