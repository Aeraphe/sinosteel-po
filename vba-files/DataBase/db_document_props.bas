Attribute VB_Name = "db_document_props"


'namespace=vba-files\DataBase



Public Function getAll(doc_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT  PT.id, PT.name, DPRO.value FROM  document_properties AS DPRO INNER JOIN  document_properties_types As PT ON PT.id = DPRO.property_id   Where  DPRO.document_id  = " & doc_id & "  ORDER BY DPRO.create_ate DESC "


    Set getAll = database.cn.Execute(sqlStrQuery)


End Function




'/*
'
'
'Search Document property type
'
'*/
Public Function SearchType(ByVal prop_name As String, Optional ByVal limit As Long = 1) As Variant

    Dim sqlStrQuery As String
    Dim database As Object
    Set database = XdbFactory.Create
    sqlStrQuery = "SELECT  * FROM  document_properties_types  Where  name='" & prop_name & "'  ORDER BY id DESC  LIMIT " & limit
    Set SearchType = database.cn.Execute(sqlStrQuery)

End Function

'/*
'Delete
'
'*/
Public Function delete(ByVal doc_id As String, ByVal prop_id As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  document_properties  Where document_id = " & doc_id & " and  property_id=" & prop_id

    database.cn.Execute (sqlStrQuery)



End Function


'/*
'
'Create
'
'@param <Array>  data
'
'*/
Public Function Create(ByRef data As Variant)

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Call database.Insert("document_properties", data)

End Function
