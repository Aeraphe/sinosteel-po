Attribute VB_Name = "Shared_PurchaseOrderSelectComp"


'namespace=vba-files\Shared


Function Mount(ByRef selecComp As Object, ByVal project_id As String)

    Dim respQuery As ADODB.Recordset
    Set respQuery = PurchaseDataBase.getAllProjectPurchase(project_id)

    selecComp.Clear

    Do Until respQuery.EOF

        selecComp.AddItem XdbFactory.getData(respQuery, "id")
        selecComp.List(selecComp.ListCount - 1, 1) = XdbFactory.getData(respQuery, "description")
        selecComp.List(selecComp.ListCount - 1, 2) = XdbFactory.getData(respQuery, "po_code")


        respQuery.MoveNext
    Loop

End Function
