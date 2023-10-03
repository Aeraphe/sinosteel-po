Attribute VB_Name = "Shared_SuppliersSelectComp"


'namespace=vba-files\Shared


Function Mount(ByRef selecComp As Object)

    Dim respQuery As ADODB.Recordset
    Set respQuery = SuppliersDataBase.getAll()

    selecComp.Clear

    Do Until respQuery.EOF

        selecComp.AddItem XdbFactory.getData(respQuery, "id")
        selecComp.List(selecComp.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

        respQuery.MoveNext
    Loop

End Function



Public Function GetSupplierSelectedInForm() As Object
    search_supplier_form.Show
    If (search_supplier_form.suppliers_listbox.Value <> "") Then
        Dim response As Object
        Set response = CreateObject("Scripting.Dictionary")


        response("id") = search_supplier_form.suppliers_listbox.Value
        x = search_supplier_form.suppliers_listbox.ListIndex
        response("name") = search_supplier_form.suppliers_listbox.List(x, 1)
        Unload search_supplier_form


        Set GetSupplierSelectedInForm = response
    End If
End Function
