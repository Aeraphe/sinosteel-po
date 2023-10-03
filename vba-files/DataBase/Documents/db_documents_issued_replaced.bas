Attribute VB_Name = "db_documents_issued_replaced"


'namespace=vba-files\DataBase\Documents



Public Function Create(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create
    On Error Resume Next
    Create = database.Insert("documents_issued_replaced", data)

End Function
