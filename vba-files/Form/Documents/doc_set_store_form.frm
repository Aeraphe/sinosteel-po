VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_set_store_form 
   Caption         =   "Definir locais de para Salvar o Arquivos"
   ClientHeight    =   12465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16185
   OleObjectBlob   =   "doc_set_store_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_set_store_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents


Private selectedProjectId
Private selectedDocumentId

Private Sub search_project_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")

   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm
   On Error Resume Next
   If (data("id") <> "") Then
      project_txt.Value = data("name")
      selectedProjectId = data("id")
      Call get_project_folders_handler(selectedProjectId)
   End If
End Sub


Private Function get_project_folders_handler(ByVal project_id As String)

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_doc_store.getProjectDocumentFolders(project_id)

   folder_select.Clear
   Do Until respQuery.EOF

      folder_select.AddItem XdbFactory.getData(respQuery, "id")
      folder_select.List(folder_select.ListCount - 1, 1) = XdbFactory.getData(respQuery, "value")
      folder_select.List(folder_select.ListCount - 1, 2) = XdbFactory.getData(respQuery, "Description")


      respQuery.MoveNext
   Loop


End Function

Private Sub search_doc_btn_Click()
   Call SearchDocumentHandler("name")
End Sub


Private Sub search_doc_sinosteel_btn_Click()
   Call SearchDocumentHandler("sinosteel_doc_number")
End Sub


Private Sub search_doc_supplier_btn_Click()
   Call SearchDocumentHandler("doc_number")
End Sub



Private Function SearchDocumentHandler(tb As String)


   Dim respQuery As ADODB.Recordset
   Set respQuery = db_documents.search(selectedProjectId, doc_txt.Value, tb)

   doc_list.Clear
   Do Until respQuery.EOF

      doc_list.AddItem XdbFactory.getData(respQuery, "id")

      doc_list.List(doc_list.ListCount - 1, 1) = XdbFactory.getData(respQuery, "doc_number")
      doc_list.List(doc_list.ListCount - 1, 2) = XdbFactory.getData(respQuery, "name")
      doc_list.List(doc_list.ListCount - 1, 3) = XdbFactory.getData(respQuery, "category")
      doc_list.List(doc_list.ListCount - 1, 4) = XdbFactory.getData(respQuery, "doc_type")


      respQuery.MoveNext
   Loop

End Function
