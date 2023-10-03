VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} search_supplier_form 
   Caption         =   "Pesquisar Fornecedor"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11145
   OleObjectBlob   =   "search_supplier_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "search_supplier_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Search



Private Sub search_btn_Click()

   Call search(search_txt.Value)

End Sub


Function search(search_string As String)

   If (search_string <> "") Then

      Dim respQuery As ADODB.Recordset
      Set respQuery = ProjectDataBase.getAll(search_txt)

      Call Shared_CommonSelectComp.Mount(projects_list, respQuery)
   End If
End Function


Private Sub UserForm_Activate()
   Call Shared_SuppliersSelectComp.Mount(suppliers_listbox)
End Sub





Private Sub select_btn_Click()
   If (suppliers_listbox.Value <> "") Then
      Me.Hide
   End If
End Sub
