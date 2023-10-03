Attribute VB_Name = "DrawingsDataBase"


'namespace=vba-files\DataBase

'/*
'
'Add Drawings
'
'@param <Array>  data
'
'*/
Public Function addDrawing(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "INSERT INTO drawings ( code,rev,tag,name,description,weight, create_ate) VALUES ('" _
    & data("code") & "','" & data("rev") & "','" & data("tag") & "','" & data("name") & "','" & data("description") & "','" & data("weight") & "','" & format(Date, "yyyy-dd-mm") & "')"

    'Insert New Budget
    database.cn.Execute (sqlStrQuery)

End Function




Public Function getAll(Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "supplier_number") As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    If (searchWord = "") Then
        sqlStrQuery = "SELECT  * FROM  drawings   ORDER BY id DESC "
    Else

        sqlStrQuery = "SELECT  * FROM  drawings    Where  " & filterType & "   LIKE '%" & searchWord & "%'  ORDER BY id DESC "
    End If

    Set getAll = database.cn.Execute(sqlStrQuery)


End Function
