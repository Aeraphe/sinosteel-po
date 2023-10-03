Attribute VB_Name = "helper_zip"


'namespace=vba-files\Helpers


Function zip_file()
' With your list of Folders located in Sheet1
' in the Range A1:A3
    
    Dim FileNameZip, foldername
    Dim strDate As String, DefPath As String
    Dim oApp As Object
    Dim Fold As Range


    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    
    strDate = format(Now, " dd-mmm-yy h-mm-ss")
    FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"


    'Create empty Zip File
    NewZip (FileNameZip)


    Set oApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(foldername).items


    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.Namespace(FileNameZip).items.count = _
       oApp.Namespace(foldername).items.count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0


    MsgBox "You find the zipfile here: " & FileNameZip



End Function

