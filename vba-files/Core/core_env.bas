Attribute VB_Name = "core_env"




'namespace=vba-files\Core


Public props As Object


Function read_config_file() As Object

   Set props = CreateObject("Scripting.Dictionary")
   Dim strFile As String
   Dim config_path As String
   Dim file_lines_on_arr As Variant

   config_path = ThisWorkbook.path

   If (config_path <> "") Then

      strFile = config_path & "\" & "config.txt"
       On Error GoTo error_handler
      file_lines_on_arr = Read_UTF_8_Text_File(strFile)
  
        For Each item In file_lines_on_arr

      prop_splited = Split(item, ":", 2)
      prop = helper_string.RemoveLineBreak(UCase(prop_splited(0)))
      Value = helper_string.RemoveLineBreak(UCase(prop_splited(1)))
      props(prop) = Value
   Next item
   
 Set read_config_file = props

Exit Function
   End If

error_handler:
   MsgBox "Não foi encontrado o arquivo de configurações (config.txt)", , "Error"
   
End Function


Public Function Read_UTF_8_Text_File(file_path As String) As Variant
   'ensure reference is set to Microsoft ActiveX DataObjects library (the latest version of it).
   'under "tools/references"... references travel with the excel file, so once added, no need to worry.
   'if not you will get a type mismatch / library error on line below.

   Dim adoStream As ADODB.Stream
   Dim var_String As Variant

   Set adoStream = New ADODB.Stream


   adoStream.Charset = "UTF-8"
   adoStream.Open
   adoStream.LoadFromFile file_path 'change this to point to your text file

   var_String = Split(adoStream.ReadText, vbCrLf) 'split entire file into array - lines delimited by CRLF

   Read_UTF_8_Text_File = var_String

End Function


