Attribute VB_Name = "shared_select_grd_to_comp"


'namespace=vba-files\Shared


Function Mount(ByRef selecComp As Object)


    Dim respQuery As ADODB.Recordset
    Set respQuery = db_grd.getAll()

    selecComp.Clear

    Do Until respQuery.EOF

        selecComp.AddItem XdbFactory.getData(respQuery, "id")
        selecComp.List(selecComp.ListCount - 1, 1) = XdbFactory.getData(respQuery, "project_id")
        selecComp.List(selecComp.ListCount - 1, 2) = XdbFactory.getData(respQuery, "name")
      

        respQuery.MoveNext
    Loop

End Function
