Attribute VB_Name = "Shared_InvoiceTypeSelect"


'namespace=vba-files\Shared


Function Mount(ByRef selecComp As Object)

    Dim respQuery As ADODB.Recordset
    Set respQuery = InvoiceDataBase.getAllInvoiceTypes()

    selecComp.Clear

    Do Until respQuery.EOF

        selecComp.AddItem XdbFactory.getData(respQuery, "name")
        selecComp.List(selecComp.ListCount - 1, 1) = XdbFactory.getData(respQuery, "description")

        respQuery.MoveNext
    Loop

End Function
