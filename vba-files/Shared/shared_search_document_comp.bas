Attribute VB_Name = "shared_search_document_comp"


'namespace=vba-files\Shared


Function Mount(ByRef selecComp As Object)

    Dim respQuery As ADODB.Recordset
    Set respQuery = ProjectDataBase.getAll()

    selecComp.Clear

    Do Until respQuery.EOF

        selecComp.AddItem XdbFactory.getData(respQuery, "id")
        selecComp.List(selecComp.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

        respQuery.MoveNext
    Loop

End Function



Public Function GetSelectedInForm() As Object
    search_document_form.Show
    If (search_document_form.document_list.Value <> "") Then
        Dim response As Object
        Set response = CreateObject("Scripting.Dictionary")


        response("id") = search_document_form.document_list.Value
        x = search_document_form.document_list.ListIndex
        response("name") = search_document_form.document_list.List(x, 1)
        Unload search_document_form


        Set GetSelectedInForm = response
    End If
End Function
