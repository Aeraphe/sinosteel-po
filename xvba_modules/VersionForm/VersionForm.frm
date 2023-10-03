VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VersionForm 
   Caption         =   "Informações sobre a Planilha"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5370
   OleObjectBlob   =   "VersionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VersionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=xvba_modules\VersionForm

Private Sub btnInfoVersionClose_Click()

  Unload VersionForm

End Sub


Private Sub to_me_lb_Click()
ActiveWorkbook.FollowHyperlink Address:="mailto:alberto.aeraph@gmail.com", NewWindow:=True
Unload Me
End Sub
