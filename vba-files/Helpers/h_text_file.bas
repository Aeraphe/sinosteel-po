Attribute VB_Name = "h_text_file"



'namespace=vba-files\Helpers



Function Create(ByVal message As String, Optional ByVal fileName As String = "log.txt")

    Dim logContent As String
    Dim fso As Object
    Dim objFile As Object

    Dim logFilePath As String
    
    logFilePath = getFolderPath(fileName)


    'create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    'check if the log file already exists in the directory
    If Not fso.FileExists(logFilePath) Then
        'if log file does not exist then create one
        Set objFile = fso.CreateTextFile(logFilePath)
    Else
        'if log file already exists then open it in append mode
        Set objFile = fso.OpenTextFile(logFilePath, 8)
    End If

    'Start writing log content to the file
    logContent = message & vbCrLf  'write whatever content you need
    objFile.Write logContent

    'Close the file
    objFile.Close

    'release the object references
    Set objFile = Nothing
    Set fso = Nothing

End Function


Public Function getFolderPath(ByVal fileName As String) As String

    Dim logPath As String
    Dim objFolders As Object
    Set objFolders = CreateObject("WScript.Shell").SpecialFolders
    
   

    logPath = objFolders("desktop") & "\"

    getFolderPath = logPath & fileName



End Function
