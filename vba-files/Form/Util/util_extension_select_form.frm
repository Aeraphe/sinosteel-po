VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} util_extension_select_form 
   Caption         =   "Selecione a Extenção do Arquivo"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "util_extension_select_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "util_extension_select_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Util


Private Sub select_btn_Click()
 Me.Hide
 
End Sub

Private Sub UserForm_Activate()

   getDocExtensionstHandler

End Sub





Private Function getDocExtensionstHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllExtensions()

   extension_select.Clear
   Call Shared_CommonSelectComp.Mount(extension_select, respQuery)
End Function
