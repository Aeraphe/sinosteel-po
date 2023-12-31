VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Xlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'namespace=xvba_modules\Xlog



'/*
'
'Log Messages
'@author Alberto Eduardo
'@since 21/04/2020
'
'*/

Private param As String

Private Sub Class_Initialize()
  param = Application.ThisWorkbook.path & "\app.log"
End Sub


'/*
'
'Set Log Message
'
'@param[Integer] messageType
'@param["String"] messageText
'@param[String] filePath
'*/
Public Function message(messageType As Integer, messageText As String, Optional filePath)
 
  If IsMissing(filePath) Then
    filePath = param
  End If
  Dim fileNum As Integer
  
  fileNum = FreeFile
  
  Dim PREFIX As String
  
  PREFIX = Now & " - "
  
  
  Open filePath For Append As #fileNum
  
  Select Case messageType
   
    Case 0 'Error Message
    Print #fileNum, PREFIX & "Error:" & messageText
    Case 1 'Success
    Print #fileNum, PREFIX & "DEBUG:" & messageText
    Case Else 'No Type Set
    Print #fileNum, PREFIX & "INFO:" & messageText
  End Select
     
  Close #fileNum
  
  
  
End Function
  
  
  
  
Public Function ErrorHandller(filePath As String, macro As String)
   
  Dim messageLog As String
  messageLog = Err.description & "  [" & Err.Number & "] Macro:" & macro
   
  Call LogMessages(0, filePath, messageLog)
  ThisWorkbook.Close SaveChanges:=False
  
End Function
  
