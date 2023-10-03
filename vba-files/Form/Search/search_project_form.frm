VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} search_project_form 
   Caption         =   "Pesquisar Projeto"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   OleObjectBlob   =   "search_project_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "search_project_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Search




Private Sub projects_list_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   Me.Hide
End Sub

Private Sub UserForm_Activate()
   Call fill_project_list_box_handler("")
End Sub

Private Sub search_btn_Click()

   Call search(search_txt.Value)

End Sub



Function search(search_string As String)

   If (search_string <> "") Then
      Call fill_project_list_box_handler(search_txt)

   End If
End Function


Private Function fill_project_list_box_handler(search_txt As String)
   Dim respQuery As ADODB.Recordset
   Set respQuery = ProjectDataBase.getAll(search_txt)

   Call Shared_CommonSelectComp.Mount(projects_list, respQuery)
End Function

Private Sub select_btn_Click()
   Me.Hide
End Sub
