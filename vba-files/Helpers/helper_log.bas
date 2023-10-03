Attribute VB_Name = "helper_log"



'namespace=vba-files\Helpers

Function DebugApp(ByVal message As String)
    On Error GoTo ErrorHandler
    If (Constants.DEBUG_APP) Then
        Call createDebugLogFile(Trim(UCase(message)) & vbNewLine, Constants.DEBUG_FILE_NAME)
    End If


    Exit Function
ErrorHandler:

    MsgBox "Erro ao criar o arquivo de log: " & Err.description, vbCritical, "Error"


End Function


Private Function createDebugLogFile(ByVal message As String, Optional ByVal fileName As String = "app_log.txt")

    Dim logContent As String
    Dim fso As Object
    Dim objFile As Object

    Dim foldername As String
    ' foldername = Environ("temp") & "\sinosteel_app"
    foldername = GetDesktopPath() & "\" & Constants.FOLDER_DEBUG

    'Create folder if not Exists
    file_helper.CreateFolder foldername

    logFilePath = foldername & "\" & fileName

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
    logContent = Now() & ": " & message & vbCrLf 'write whatever content you need
    objFile.Write logContent

    'Close the file
    objFile.Close

    'release the object references
    Set objFile = Nothing
    Set fso = Nothing

End Function

Function createLogFile(ByVal message As String, Optional ByVal fileName As String = "log.txt")

    Dim logContent As String
    Dim fso As Object
    Dim objFile As Object

    Dim logFilePath As String

    logFilePath = getLogFilePath(fileName)


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
    logContent = Now() & ": " & message & vbCrLf 'write whatever content you need
    objFile.Write logContent

    'Close the file
    objFile.Close

    'release the object references
    Set objFile = Nothing
    Set fso = Nothing

End Function


Public Function getLogFilePath(ByVal fileName As String) As String

    Dim logPath As String


    Dim objFolders As Object
    Set objFolders = CreateObject("WScript.Shell").SpecialFolders


    'set the path for the log file
    logPath = objFolders("desktop") & "\"
    'set the filename for the log file

    getLogFilePath = logPath & fileName



End Function


Function GetDesktopPath() As String
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")

    ' Retrieve the path to the Desktop folder for the current user
    GetDesktopPath = WshShell.SpecialFolders("Desktop")
End Function
