Attribute VB_Name = "db_doc_flow"


'namespace=vba-files\DataBase

'/*
'
'Create
'
'@param <Array>  data
'
'*/
Public Function Create(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    Create = database.Insert("doc_flow", data)

End Function



'/*
'
'Create
'
'@param <Array>  data
'
'*/
Public Function InsterItems(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    Create = database.Insert("doc_flow_itens", data)

End Function



'
'Delete budget By ID
'
'*/
Public Function delete(ByRef idData As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  doc_flow  Where id = " & idData

    database.cn.Execute (sqlStrQuery)



End Function



'
'Delete budget By ID
'
'*/
Public Function delete_items(ByRef idData As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  doc_flow_itens  Where id = " & idData

    database.cn.Execute (sqlStrQuery)



End Function
