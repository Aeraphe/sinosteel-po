Attribute VB_Name = "db_documents_equipaments"


'namespace=vba-files\DataBase



Public Function getAll(doc_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT  EQ.id, EQ.name, EQ.type FROM  documents_equipaments AS DOC_EQ INNER JOIN  projects_equipaments As EQ ON EQ.id = DOC_EQ.equipament_id   Where  DOC_EQ.document_id  = " & doc_id & "  ORDER BY DOC_EQ.create_ate DESC "


    Set getAll = database.cn.Execute(sqlStrQuery)


End Function


'/*
'Delete
'
'*/
Public Function delete(ByVal doc_id As String, ByVal equip_id As String)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  documents_equipaments  Where document_id = " & doc_id & " and  equipament_id=" & equip_id

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

    Call database.Insert("documents_equipaments", data)

End Function
