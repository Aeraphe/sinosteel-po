Attribute VB_Name = "database_doc_proprty_types"


'namespace=vba-files\DataBase




Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("document_properties_types")

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

    sqlStrQuery = " DELETE   FROM  document_properties_types  Where id = " & idData

    database.cn.Execute (sqlStrQuery)



End Function
