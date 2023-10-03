VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} util_doc_code_select_form 
   Caption         =   "Selecione o código do documento"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "util_doc_code_select_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "util_doc_code_select_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Util



Private Sub UserForm_Activate()
 getDocCodeTypestHandler
End Sub

Private Function getDocCodeTypestHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.getAllDocCodes()

   doc_code_seletc.Clear
   Call Shared_CommonSelectComp.Mount(doc_code_seletc, respQuery)
End Function
Private Sub select_btn_Click()
 Me.Hide
 
End Sub
