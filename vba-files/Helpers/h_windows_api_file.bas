Attribute VB_Name = "h_windows_api_file"




'namespace=vba-files\Helpers

Option Explicit


Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32.dll" Alias "SHCreateDirectoryExA" (ByVal hWnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As LongPtr
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As LongPtr
Private Declare PtrSafe Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As LongPtr

Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As LongPtr, ByVal dwShareMode As LongPtr, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As LongPtr, ByVal dwFlagsAndAttributes As LongPtr, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As LongPtr

Private Const GENERIC_READ As LongPtr = &H80000000
Private Const OPEN_EXISTING As LongPtr = 3&
Private Const FILE_FLAG_SEQUENTIAL_SCAN As LongPtr = &H8000000
Private Const PROV_RSA_FULL As LongPtr = 1&
Private Const CALG_MD5 As LongPtr = 32771

    Private Type CRYPTPROV
    cbSize As Long
    dwProvType As Long
    pbProvData As Long
    dwFlags As Long
    dwProvType2 As Long
    pbProvData2 As Long
    dwReserved2 As Long
    End Type

    Private Type CRYPTKEY
    cbSize As Long
    dwSessionKey As Long
    dwFlags As Long
    End Type

    Private Type SHFILEOPSTRUCT
    hWnd As LongPtr
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As LongPtr
    lpszProgressTitle As String
    End Type
Function CopyFileLongPathWithChecksum(SourceFilePath As String, destinationFolderPath As String, Optional deleteOriginalFile As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileOp As SHFILEOPSTRUCT
    Dim retryCount As Integer
    Dim originalChecksum As String
    Dim copiedChecksum As String
    Dim result As LongPtr
    
    ' Set up the SHFILEOPSTRUCT parameters
    fileOp.hWnd = 1 ' Set hWnd to 0 to hide the copy process window
    fileOp.wFunc = 2 ' FO_COPY: Specifies a file copy operation
    fileOp.pFrom = SourceFilePath & vbNullChar ' Source file path
    fileOp.pTo = destinationFolderPath & vbNullChar ' Destination folder path
    fileOp.fFlags = &H10 Or &H200 Or &H400
    ' FOF_NOCONFIRMATION: Do not display a confirmation dialog for file operations
    ' FOF_SILENT: Do not display any dialogs
    
    retryCount = 0
    
    Do While retryCount < 5
    
   'Call ForceCloseFile(SourceFilePath)
        ' Perform the file copy operation using the Windows API
        result = SHFileOperation(fileOp)
        
        If result = 0& And Not fileOp.fAnyOperationsAborted Then
            originalChecksum = h_check2.CalculateFileChecksum(SourceFilePath)
            copiedChecksum = h_check2.CalculateFileChecksum(destinationFolderPath & "\" & GetFileNameFromPath(SourceFilePath))
            
            If originalChecksum = copiedChecksum Then
                If deleteOriginalFile Then
                    DeleteFile SourceFilePath
                End If
                
                CopyFileLongPathWithChecksum = True
                Exit Function
            End If
        End If
        
        retryCount = retryCount + 1
    Loop
    
    CopyFileLongPathWithChecksum = False
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description
    CopyFileLongPathWithChecksum = False
End Function



Function GetFileNameFromPath(filePath As String) As String
    On Error Resume Next
    GetFileNameFromPath = VBA.Strings.Mid(filePath, InStrRev(filePath, "\") + 1)
End Function

Function CreateFolderHierarchy(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Call the SHCreateDirectoryEx function to create the folder hierarchy
    Dim result As LongPtr
    result = SHCreateDirectoryEx(0, folderPath, 0)
    
    ' Check if the folder hierarchy was created successfully
    If result = 0& Then
        CreateFolderHierarchy = False ' Return False if folder hierarchy creation failed
    Else
        CreateFolderHierarchy = True ' Return True if folder hierarchy was created successfully
    End If
    
    Exit Function
    
ErrorHandler:
    CreateFolderHierarchy = False ' Return False if an error occurred
End Function




Sub ForceCloseFile(ByVal filePath As String)
    On Error GoTo ErrorHandler
    
    Dim hFile As LongPtr
    
    ' Open the file
    hFile = CreateFile(filePath, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    
    If hFile <> 0& Then
        ' Close the file handle
        CloseHandle hFile
        
        MsgBox "File closed successfully."
    Else
        MsgBox "Failed to open the file."
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description
End Sub
