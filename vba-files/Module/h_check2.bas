Attribute VB_Name = "h_check2"
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As LongPtr, ByVal dwShareMode As LongPtr, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As LongPtr, ByVal dwFlagsAndAttributes As LongPtr, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpFileSizeHigh As LongPtr) As LongPtr
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As LongPtr, ByRef lpNumberOfBytesRead As LongPtr, ByVal lpOverlapped As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32.dll" (ByRef phProv As LongPtr, ByVal pszContainer As LongPtr, ByVal pszProvider As LongPtr, ByVal dwProvType As LongPtr, ByVal dwFlags As LongPtr) As LongPtr
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal Algid As LongPtr, ByVal hKey As LongPtr, ByVal dwFlags As LongPtr, ByRef phHash As LongPtr) As LongPtr
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hHash As LongPtr, ByRef pbData As Any, ByVal dwDataLen As LongPtr, ByVal dwFlags As LongPtr) As LongPtr
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As LongPtr, ByVal dwParam As LongPtr, ByRef pbData As Any, ByRef pdwDataLen As LongPtr, ByVal dwFlags As LongPtr) As LongPtr
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As LongPtr) As LongPtr
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As LongPtr) As LongPtr

Private Const PROV_RSA_FULL As LongPtr = 1
Private Const CALG_MD5 As LongPtr = 32771
Private Const HP_HASHVAL As LongPtr = 2
Private Const HP_HASHSIZE As LongPtr = 4

Function CalculateFileChecksum(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim hFile As LongPtr
    Dim fileSize As LongPtr
    Dim buffer() As Byte
    Dim bytesRead As LongPtr
    Dim result As LongPtr
    Dim hProv As LongPtr
    Dim hHash As LongPtr
    Dim hashSize As LongPtr
    Dim hashValue() As Byte
    Dim i As Long
    
    ' Convert the file path to a wide string
    Dim filePathW As LongPtr
    filePathW = StrPtr(filePath)
    
    ' Open the file
    hFile = CreateFileW(filePathW, &H80000000, 0, 0, 3, 0, 0)
    
    If hFile <> 0 Then
        ' Get the file size
        fileSize = GetFileSize(hFile, 0)
        
           
              Dim bufferSize As Long
            bufferSize = CLng(fileSize)
            
        ' Allocate the buffer
        ReDim buffer(0 To bufferSize - 1) As Byte
        
        ' Read the file content into the buffer
        result = ReadFile(hFile, buffer(0), fileSize, bytesRead, 0)
        
        If result <> 0 Then
            ' Acquire a cryptographic provider context
            result = CryptAcquireContextW(hProv, 0, 0, PROV_RSA_FULL, 0)
            
            If result <> 0 Then
                ' Create a hash object
                result = CryptCreateHash(hProv, CALG_MD5, 0, 0, hHash)
                
                If result <> 0 Then
                    ' Hash the file content
                    result = CryptHashData(hHash, buffer(0), bytesRead, 0)
                    
                    If result <> 0 Then
                        ' Get the hash size
                        result = CryptGetHashParam(hHash, HP_HASHSIZE, hashSize, 4, 0)
                        
                        If result <> 0 Then
                            ' Allocate the hash value buffer
                            ReDim hashValue(0 To CLng(hashSize) - 1) As Byte
                            
                            ' Get the hash value
                            result = CryptGetHashParam(hHash, HP_HASHVAL, hashValue(0), hashSize, 0)
                            
                            If result <> 0 Then
                                ' Convert the hash value to a hex string
                                Dim checksum As String
                                For i = 0 To UBound(hashValue)
                                    checksum = checksum & Right$("0" & Hex$(hashValue(i)), 2)
                                Next i
                                
                                CalculateFileChecksum = checksum
                            End If
                        End If
                    End If
                    
                    ' Destroy the hash object
                    CryptDestroyHash hHash
                End If
                
                ' Release the cryptographic provider context
                CryptReleaseContext hProv, 0
            End If
        End If
        
        ' Close the file handle
        CloseHandle hFile
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.description
    CalculateFileChecksum = ""
End Function

