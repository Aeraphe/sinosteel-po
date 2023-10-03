Attribute VB_Name = "action_organize_project_files"

'namespace=vba-files\Actions



Sub start(ByVal projectSelectedId As String, ByVal folderPath As String, Optional ByVal folderType As String = "SENT")
    Dim filesDict As Object

    ' Initialization
    Set filesDict = InitializeFileMove(folderPath)

    ' Process each file
    ProcessFiles filesDict, projectSelectedId, folderType

    ' Close the alert form
    Unload UserFormAlert
 Exit Sub

ErrorHandler:
    HandleError
End Sub

Private Function InitializeFileMove(ByVal folderPath As String) As Object
    Dim filesDict As Object
    Set filesDict = file_helper.get_files_from_folders2(folderPath)

    ' Show user form alert
    Load UserFormAlert
    UserFormAlert.Label1.Caption = "Movendo os arquivos"
    UserFormAlert.Show

    Set InitializeFileMove = filesDict
End Function

Private Sub ProcessFiles(ByRef filesDict As Object, ByVal projectSelectedId As String, Optional ByVal folderType As String = "SENT")
    Dim varKey As Variant
    For Each varKey In filesDict.Keys()
        If varKey <> "count" Then
            MoveAndLogFile filesDict(varKey), projectSelectedId, folderType
        End If
    Next varKey
End Sub



Private Sub MoveAndLogFile(ByRef fileInfo As Dictionary, ByVal projectSelectedId As String, Optional ByVal folderType As String = "SENT")
    On Error GoTo ErrorHandler

        Dim fileNameSplitted() As String
        Dim doc_number As String
        Dim originFilePath As String
        Dim respQuery As ADODB.Recordset

        Dim searchFileName As String, fileExtension As String




        searchFileName = fileInfo("file")
        fileExtension = fileInfo("extension")
        originFilePath = fileInfo("path")

        ' Extract doc number from the file name
        If InStr(UCase(searchFileName), "_REV_") Then
            doc_number = Split(UCase(searchFileName), "_REV_")(0)
        Else
            doc_number = Left(searchFileName, Len(searchFileName) - InStrRev(searchFileName, "_") - 1)
        End If

        ' Query the database For document details
        Set respQuery = db_documents.SearchLimit(projectSelectedId, doc_number, "doc_number")

        Dim docId As String, contractItem As String, disciplineID As String
        docId = XdbFactory.getData(respQuery, "id")
        contractItem = XdbFactory.getData(respQuery, "contract_item")
        disciplineID = XdbFactory.getData(respQuery, "discipline_id")

        ' Check If required document details are present
        If (docId <> "" And contractItem <> "" And disciplineID <> "") Then
            UserFormAlert.labelInfo.Caption = "Pasta: " & contractItem & _
            "  Arquivos Encontrados:  [ " & docCount & " ] " & _
            "  Arquivos Movidos: [ " & fileCount & " ]"
            UserFormAlert.Repaint

            ' Move the file And wait For the operation To complete
            Call shared_project_files.moveFile(originFilePath, projectSelectedId, docId, searchFileName & "." & fileExtension, folderType)
            Call Xhelper.waitMs(800)
        End If

     Exit Sub
ErrorHandler:
        HandleError
End Sub
   

Private Sub HandleError()
    MsgBox "An error occurred: " & Err.description, vbExclamation, "Error"
    ' Optionally: log the error Or perform additional error handling here
End Sub
