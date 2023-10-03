Attribute VB_Name = "SuppliersDataBase"


'namespace=vba-files\DataBase

'/*
'
'Create Sppliers
'
'@param <Array>  data
'
'*/
Public Function Create(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "INSERT INTO suppliers ( name,type,email,phone,address, create_ate) VALUES ('" _
    & data("name") & "','" & data("type") & "','" & data("email") & "','" & data("phone") & "','" & data("address") & "','" & format(Date, "yyyy-dd-mm") & "')"

    'Insert New Budget
    database.cn.Execute (sqlStrQuery)

End Function




Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("suppliers")

End Function




Public Function getAllSuppliersFromDrawing(drawingId As String) As ADODB.Recordset



    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    sqlStrQuery = "SELECT  sup.name as name, sup.id as id FROM  drawing_manufactures As dsup  INNER JOIN suppliers As sup ON sup.id=dsup.manufacturer_id   Where drawing_id=" & drawingId & "  ORDER BY dsup.create_ate DESC "


    Set getAllSuppliersFromDrawing = database.cn.Execute(sqlStrQuery)


End Function



'/*
'
'Create Sppliers
'
'@param <Array>  data
'
'*/
Public Function SetManufactureToDrawing(ByRef data As Variant)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "INSERT INTO drawing_manufactures ( drawing_id,manufacturer_id, create_ate) VALUES ('" _
    & data("drawing_id") & "','" & data("manufactor_id") & "','" & format(Date, "yyyy-dd-mm") & "')"

    'Insert New Budget
    database.cn.Execute (sqlStrQuery)

End Function


'/*
'
'Delete budget By ID
'
'*/
Public Function RemoveMAnufacturerFromDrawing(ByRef data As Variant) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  drawing_manufactures  Where drawing_id = " & data("drawing_id") & " AND  manufacturer_id = " & data("manufactor_id")

    'Set the recor set
    Set rs = New ADODB.Recordset
    Set rs = database.cn.Execute(sqlStrQuery)

    Set RemoveMAnufacturerFromDrawing = rs


End Function
