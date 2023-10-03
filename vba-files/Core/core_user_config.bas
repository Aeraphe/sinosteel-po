Attribute VB_Name = "core_user_config"

'namespace=vba-files\Core


Option Explicit

' Declare the user configurations dictionary as a public variable
Public userConfigurations As Object
Private userConfigFile As String


Sub SetDefaultConfigurations()

    Set userConfigurations = CreateObject("Scripting.Dictionary")
    ' Set default values for each configuration setting
    userConfigurations("stayLoggedIn") = True
    userConfigurations("language") = "Portugues"
    ' Add additional default configuration settings here
End Sub

Sub AddUserConfiguration(key As String, Value As Variant)
    ' Check if the key already exists in the dictionary
    If Not userConfigurations.Exists(key) Then
        ' Add the key-value pair to the dictionary
        userConfigurations.Add key, Value
    Else
        ' Update the existing key-value pair
        userConfigurations(key) = Value
    End If
End Sub

'/*
' Procedure to read user configurations from file
'*/
Sub ReadUserConfigurationsFromFile()
    On Error GoTo ErrorHandler

    ' Declare and initialize the dictionary object
    Dim userConfigurations As Object
    Set userConfigurations = CreateObject("Scripting.Dictionary")
    
    ' Declare and initialize the FSO object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists before attempting to open it
    Dim fileName As String
    fileName = Environ("temp") & "\sinosteel_app\user_configurations.txt"
    If fso.FileExists(fileName) Then
        ' Open the file for reading
        Dim file As Object
        Set file = fso.OpenTextFile(fileName)
        
        ' Loop through each line in the file
        Do Until file.AtEndOfStream
            ' Read the next line from the file
            Dim line As String
            line = file.ReadLine
            
            If InStr(line, "=") > 0 Then
                ' Split the line into key-value pairs and add them to the dictionary
                Dim keyValuePairs() As String
                keyValuePairs = Split(line, "=")
                userConfigurations(keyValuePairs(0)) = keyValuePairs(1)
            End If
        Loop
        
        ' Close the file
        file.Close
    End If
    
    ' Do something with the userConfigurations dictionary object here
    ' ...
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while attempting to read user configurations from file:" & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & _
        "Error Description: " & Err.description & vbNewLine & _
        "File Name: " & fileName
End Sub





Sub WriteUserConfigurationsToFile()
    On Error GoTo ErrorHandler
    
    ' Declare and initialize the FSO object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create the directory if it doesn't exist
    Dim foldername As String
    foldername = Environ("temp") & "\sinosteel_app"
    If Not fso.FolderExists(foldername) Then
        fso.CreateFolder foldername
    End If
    
    ' Open the file for writing (creates a new file if it doesn't exist)
    Dim fileName As String
    fileName = foldername & "\user_configurations.txt"
    Dim file As Object
    Set file = fso.CreateTextFile(fileName, True)
    
    ' Loop through each key-value pair in the dictionary and write them to the file
    Dim key As Variant
    For Each key In userConfigurations.Keys()
        file.WriteLine key & "=" & userConfigurations(key)
    Next key
    
    ' Close the file
    file.Close
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while attempting to write user configurations to file:" & vbNewLine & _
        "Error Number: " & Err.Number & vbNewLine & _
        "Error Description: " & Err.description & vbNewLine & _
        "File Name: " & fileName
End Sub



Sub UpdateStayLoggedInValue(newValue As Boolean)
    ' Update the "stay logged in" configuration setting
    On Error GoTo DefaultValues
    
    If IsObject(userConfigurations) Then
        userConfigurations("stayLoggedIn") = newValue
        WriteUserConfigurationsToFile
    End If
    Exit Sub
    
DefaultValues:
    SetDefaultConfigurations
    userConfigurations("stayLoggedIn") = newValue
    WriteUserConfigurationsToFile
End Sub
