VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_selected_info_form 
   Caption         =   "Informações do Documento Selecionado"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19170
   OleObjectBlob   =   "doc_selected_info_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_selected_info_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Dcuments_Flow





Public Function load_data(doc_review_id As String)

  If (doc_review_id <> "") Then
    Dim doc_id As String
    doc_id = get_doc_info(doc_review_id)
    Call load_doc_properties(doc_id)
  End If
End Function


Private Function get_doc_info(doc_review_id As String) As String

  Dim query_data  As Object
  Dim doc_id As String

  Set query_data = CreateObject("Scripting.Dictionary")
  query_data("PROP") = doc_review_id
  Dim respQuery As ADODB.Recordset
  Set respQuery = db_documents.get_document_by_review_id(query_data)

  doc_id = respQuery.fields.item("id").Value

  doc_number_txt.Value = respQuery.fields.item("doc_number").Value
  first_title.Value = respQuery.fields.item("name").Value
  last_title.Value = respQuery.fields.item("description").Value
  doc_review_txt.Value = respQuery.fields.item("rev_code").Value
  doc_issue_txt.Value = respQuery.fields.item("issue").Value
  doc_status_txt.Value = respQuery.fields.item("status").Value
  sinosteel_number_txt.Value = respQuery.fields.item("sinosteel_doc_number").Value

  get_doc_info = doc_id

End Function


Private Function load_doc_properties(doc_id As String)

  Dim respQuery As ADODB.Recordset
  Set respQuery = db_document_props.getAll(doc_id)


  prop_select.Clear


  Do Until respQuery.EOF

    prop_select.AddItem respQuery.fields.item("name")
    prop_select.List(prop_select.ListCount - 1, 1) = respQuery.fields.item("value")

    respQuery.MoveNext
  Loop



End Function




Private Sub btn_copy_titles_Click()
  Call Clipboard(first_title.Value & " - " & last_title.Value)
  Call Alert.Show("Copiado", "", 1000)
End Sub

Private Sub doc_number_txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call Clipboard(doc_number_txt.Value)
  Call Alert.Show("Copiado", "", 1000)

End Sub



Private Sub first_title_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim pro_value As String
  pro_value = first_title.Value
  Call Clipboard(pro_value)
  Call Alert.Show("Copiado", "", 1000)
End Sub



Private Sub last_title_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim pro_value As String
  pro_value = last_title.Value
  Call Clipboard(pro_value)
  Call Alert.Show("Copiado", "", 1000)
End Sub

Private Sub prop_select_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim pro_value As String
  On Error Resume Next
  pro_value = prop_select.List(prop_select.ListIndex, 1)
  If (pro_value <> "") Then
    Call Clipboard(pro_value)
    Call Alert.Show("Copiado", "", 1000)
  End If
End Sub


Function Clipboard(Optional StoreText As String) As String
  'PURPOSE: Read/Write to Clipboard
  'Source: ExcelHero.com (Daniel Ferry)

  Dim x As Variant

  'Store as variant for 64-bit VBA support
  x = StoreText

  'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
       Case Len(StoreText)
        'Write to the clipboard
        .setData "text", x
       Case Else
        'Read from the clipboard (no variable passed through)
        Clipboard = .getData("text")
      End Select
    End With
  End With

End Function

Private Sub TextBox2_Change()

End Sub
