VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_Open()


 ' bootstrap
   ' Set default configurations when the workbook is opened
   SetDefaultConfigurations
   ReadUserConfigurationsFromFile

End Sub



Private Function Bootstrap()

  On Error Resume Next
  Dim Mdl As Object
  On Error Resume Next
  Set Mdl = ActiveWorkbook.VBProject.VBComponents("action_bootstrap").CodeModule
  On Error GoTo 0
  If Mdl Is Nothing Then
    Exit Function
  Else
    action_bootstrap.run
  End If
End Function


