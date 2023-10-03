Attribute VB_Name = "Shared_DocCategorySelectComp"


'namespace=vba-files\Shared


Function Mount(ByRef selecComp As Object)

    Dim respQuery As ADODB.Recordset
    Set respQuery = DocCategoryDataBase.getAll()

    selecComp.Clear

    Do Until respQuery.EOF

        selecComp.AddItem XdbFactory.getData(respQuery, "id")
        selecComp.List(selecComp.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")

        respQuery.MoveNext
    Loop

End Function
