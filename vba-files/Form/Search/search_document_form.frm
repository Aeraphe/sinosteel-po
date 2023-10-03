VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} search_document_form 
   Caption         =   "Buscar Documento"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   OleObjectBlob   =   "search_document_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "search_document_form"
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
      Set respQuery = db_documents.search(search_txt)

      Call Shared_CommonSelectComp.Mount(document_list, respQuery)
   End If
End Function



Private Sub select_btn_Click()
   Me.Hide
End Sub
