Attribute VB_Name = "Shared_ProjectSelectComp"


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



Public Function GetProjetSelectedInForm() As Object
    search_project_form.Show
    If (search_project_form.projects_list.Value <> "") Then
        Dim response As Object
        Set response = CreateObject("Scripting.Dictionary")


        response("id") = search_project_form.projects_list.Value
        x = search_project_form.projects_list.ListIndex
        response("name") = search_project_form.projects_list.List(x, 1)
        Unload search_project_form


        Set GetProjetSelectedInForm = response
    End If
End Function
