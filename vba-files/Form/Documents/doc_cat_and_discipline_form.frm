VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_cat_and_discipline_form 
   Caption         =   "Adicionar/Modificar Categoria e Disciplina do Documento"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15390
   OleObjectBlob   =   "doc_cat_and_discipline_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_cat_and_discipline_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents

Public project_selected_id


Private Sub search_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      project_selected_id = data("id")


   End If
End Sub

Private Sub UserForm_Activate()
     Call Shared_DocCategorySelectComp.Mount(doc_category_select)
  getDisciplineHandler
End Sub


Private Function getDisciplineHandler()
   Dim respQuery As ADODB.Recordset
   Set respQuery = database_discipline.getAll()

   discipline_select.Clear
   Call Shared_CommonSelectComp.Mount(discipline_select, respQuery)
End Function
