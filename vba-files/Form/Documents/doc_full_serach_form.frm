VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_full_serach_form 
   Caption         =   "Pesquisa de Documentos"
   ClientHeight    =   13260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23325
   OleObjectBlob   =   "doc_full_serach_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_full_serach_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents

Public project_selected_id




Private Sub UserForm_Activate()


   Call auth.is_logged_to_access(Me)


End Sub
