VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_reject_form 
   Caption         =   "Rejeitar Docuemtnos Selecionados"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10515
   OleObjectBlob   =   "doc_reject_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_reject_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Dcuments_Flow


Public projectName As String
Public rejectDocList As Object
Public projectSelectedId As String
Public importFilesFolderPath As String


Private Sub reject_document_btn_Click()
 Dim answer As Integer
   Dim docRejectedList  As Object
   Dim Outlook_App As Object
   If (reject_motive.Value <> "") Then

   answer = MsgBox("Tem certeza que quer REJEITAR o(s) Docuemnto(s) ?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

   If (answer = vbYes) Then
   On Error Resume Next
   Set Outlook_App = GetObject(, "Outlook.Application")

   If Err.Number = 429 Then
      MsgBox "Erro: Gentileza abrir o Outlook Primeiro", vbCritical
   Else

      Set docRejectedList = reject_document_handler
      If (Not docRejectedList Is Nothing) Then
         Call act_reject_notifi.make(projectName, docRejectedList, Outlook_App)
     
      End If
      Call Alert.Show("Documento(s) Rejeitado(s) com Sucesso", "", 25000)
      Unload Me
   End If
   End If
      End If
      
End Sub


Private Function reject_document_handler() As Object

   Dim reject_id As Long
   Dim countRejected As Integer
   Dim docRequestId As String
   Dim docSelectedId As String
   Dim fileFullPath As String
   Dim doc_number_splited() As String
   Dim response As Object
   Dim docReject As Object
   Dim data As Object
   Dim rejectedDocIndex As String
   Dim f As Object
   
   Set response = CreateObject("Scripting.Dictionary")
   

     

Set fso = CreateObject("Scripting.FileSystemObject")

    countRejected = 0
    
    For Each varKey In rejectDocList.Keys()
       Set docReject = rejectDocList(varKey)
    
             On Error Resume Next
              
                Set data = CreateObject("Scripting.Dictionary")
                data("motive") = UCase(reject_motive.Value)
                data("doc_id") = docReject("docId")
                
                If (docReject("reviewId")) Then
                   data("review_id") = docReject("reviewId")
                   data("replace_file") = docReject("NEED")
                End If
             
                data("user_id") = docReject("userId")
                data("review") = docReject("review")
                data("request_doc_id") = docReject("requestDocID")
                
                fileFullPath = docReject("fileFullPath")
                data("file_name") = fso.GetFileName(fileFullPath)
                data("file_path") = fso.GetAbsolutePathName(fileFullPath)
                data("file_extension") = fso.GetExtensionName(fileFullPath)
                Set f = fso.GetFile(fileFullPath)
                data("file_size") = f.Size
                          
                data("FOLDER") = action_reject_document.cdoc_reject(data, projectSelectedId, docReject("fileFullPath"))
                
                'Get Request Doc Id for Change the Status to Rejected
                Call act_doc_request.changeDocRequestStatus(docReject("requestDocID"), "REJEITADO")
               
                rejectedDocIndex = "n" & countRejected
                response(rejectedDocIndex) = data
            
                
                
                countRejected = countRejected + 1
                
                  

           Call helper_log.createLogFile("ARQUIVO REJEITADO: " & docReject("docNumber") & " REV: " & docReject("review") & "  -  " & reject_motive.Value, "rejeitados.log")
        
          Next


Set reject_document_handler = response
End Function

Private Sub UserForm_Click()

End Sub
