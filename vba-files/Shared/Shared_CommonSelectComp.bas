Attribute VB_Name = "Shared_CommonSelectComp"

'namespace=vba-files\Shared


Function Mount(ByRef selecComp As Object, respQuery As ADODB.Recordset, Optional ByVal field1 As String = "id", Optional ByVal field2 As String = "name")


    selecComp.Clear

    Do Until respQuery.EOF

        selecComp.AddItem XdbFactory.getData(respQuery, field1)
        selecComp.List(selecComp.ListCount - 1, 1) = XdbFactory.getData(respQuery, field2)

        respQuery.MoveNext
    Loop

End Function

