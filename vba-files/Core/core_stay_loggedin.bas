Attribute VB_Name = "core_stay_loggedin"

Option Explicit

'namespace=vba-files\Core


Sub SaveCredentials(login As String, password As String)
    ' Saves the login and password to a text file with encryption

    Dim Shift As Integer
    Dim filePath As String
    Dim msg As String

    ' Set the encryption shift value (e.g. 3)
    Shift = 3

    ' Set the file path for storing the encrypted credentials
    filePath = Environ("temp") & "\MyCredentials.txt"

    ' Encrypt and save the login and password to the text file
    If EncryptAndSave(login & "|" & password, Shift, filePath) Then
     msg = "Credentials saved successfully"
    Else
       msg = "Failed to save credentials"
    End If
    
    Call helper_log.DebugApp(msg)
End Sub


Function GetCredentials() As Object
    ' Checks if the entered login and password match the saved credentials

    Dim Shift As Integer
    Dim filePath As String
    Dim savedCredentials As String
    Dim response As Object

    ' Set the encryption shift value (e.g. 3)
    Shift = 3

    ' Set the file path for retrieving the encrypted credentials
    filePath = Environ("temp") & "\MyCredentials.txt"

    ' Decrypt the saved credentials from the text file
    savedCredentials = DecryptFromFile(filePath, Shift)

    ' Split the saved credentials into login and password
    Dim savedLogin As String
    Dim savedPassword As String
    Dim loginPassword() As String

  
    
    Set response = CreateObject("Scripting.Dictionary")
    loginPassword = Split(savedCredentials, "|")
    savedLogin = loginPassword(0)
    savedPassword = loginPassword(1)
    response("savedLogin") = loginPassword(0)
    response("savedPassword") = loginPassword(1)

  Set GetCredentials = response
End Function






Private Function EncryptAndSave(plainText As String, Shift As Integer, filePath As String) As Boolean
    ' Encrypts the plain text with a Caesar cipher and saves it to a text file

    Dim i As Integer
    Dim charCode As Integer
    Dim encryptedText As String
    Dim fileNumber As Integer

    ' Open the text file for writing
    fileNumber = FreeFile()
    Open filePath For Output As #fileNumber

    ' Loop through each character in the plain text
    For i = 1 To Len(plainText)
        charCode = Asc(Mid(plainText, i, 1))

        ' Shift the character code by the specified amount
        If charCode >= 65 And charCode <= 90 Then ' Upper case A-Z
            charCode = ((charCode - 65 + Shift) Mod 26) + 65
        ElseIf charCode >= 97 And charCode <= 122 Then ' Lower case a-z
            charCode = ((charCode - 97 + Shift) Mod 26) + 97
        Else ' Non-letter characters (e.g. spaces, numbers, symbols, etc.)
            ' Do not shift the character code
        End If

        ' Append the encrypted character to the encrypted text
        encryptedText = encryptedText & Chr(charCode)
    Next i

    ' Write the encrypted text to the text file
    Print #fileNumber, encryptedText

    ' Close the text file
    Close #fileNumber

    EncryptAndSave = True
End Function


Private Function DecryptFromFile(filePath As String, Shift As Integer) As String
    ' Decrypts the encrypted text in a text file with a Caesar cipher

    Dim fileNumber As Integer
    Dim encryptedText As String
    Dim plainText As String

    ' Open the text file for reading
    fileNumber = FreeFile()
    Open filePath For Input As #fileNumber

    ' Read the encrypted text from the text file
    Line Input #fileNumber, encryptedText

    ' Close the text file
    Close #fileNumber

    Dim i As Long
    Dim charCode As Long
    ' Loop through each character in the encrypted text
    For i = 1 To Len(encryptedText)
        charCode = Asc(Mid(encryptedText, i, 1))

        ' Shift the character code in the opposite direction of encryption
        If charCode >= 65 And charCode <= 90 Then ' Upper case A-Z
            charCode = ((charCode - 65 - Shift + 26) Mod 26) + 65
        ElseIf charCode >= 97 And charCode <= 122 Then ' Lower case a-z
            charCode = ((charCode - 97 - Shift + 26) Mod 26) + 97
        Else ' Non-letter characters (e.g. spaces, numbers, symbols, etc.)
            ' Do not shift the character code
        End If

        ' Append the decrypted character to the plain text
        plainText = plainText & Chr(charCode)
    Next i

    DecryptFromFile = plainText
End Function

