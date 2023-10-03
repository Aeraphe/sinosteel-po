Attribute VB_Name = "h_ckecksum"
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As LongPtr, ByVal dwShareMode As LongPtr, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As LongPtr, ByVal dwFlagsAndAttributes As LongPtr, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpFileSizeHigh As LongPtr) As LongPtr
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As LongPtr, ByRef lpNumberOfBytesRead As LongPtr, ByVal lpOverlapped As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As LongPtr

Function CalculateFileChecksum(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim hFile As LongPtr
    Dim fileSize As LongPtr
    Dim buffer() As Byte
    Dim bytesRead As LongPtr
    Dim result As LongPtr
    
    ' Convert the file path to a wide string
    Dim filePathW As LongPtr
    filePathW = StrPtr(filePath)
    
    ' Open the file
    hFile = CreateFileW(filePathW, &H80000000, 0, 0, 3, 0, 0)
    
    If hFile <> 0& Then
        ' Get the file size
        fileSize = GetFileSize(hFile, 0&)
        
              Dim bufferSize As Long
            bufferSize = CLng(fileSize)
        
        ' Allocate the buffer
        ReDim buffer(0 To bufferSize - 1) As Byte
        
        ' Read the file content into the buffer
        result = ReadFile(hFile, buffer(0), fileSize, bytesRead, 0)
        
        If result <> 0& Then
            ' Calculate the MD5 hash
            Dim md5 As Object
            Set md5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
            Dim hashValue() As Byte
            hashValue = md5.ComputeHash(buffer)
            
            ' Convert the hash value to a hex string
            Dim i As Long
            Dim checksum As String
            For i = 0 To UBound(hashValue)
                checksum = checksum & Right("0" & Hex(hashValue(i)), 2)
            Next i
            
            CalculateFileChecksum = checksum
        End If
        
        ' Close the file handle
        CloseHandle hFile
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description
    CalculateFileChecksum = ""
End Function

