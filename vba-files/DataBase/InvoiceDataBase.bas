Attribute VB_Name = "InvoiceDataBase"


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

    Create = database.Insert("invoice", data)

End Function




'/*
'
'Create Sppliers
'
'@param <Array>  data
'
'*/
Public Function CreateTypes(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    CreateTypes = database.Insert("invoice_types", data)

End Function


Public Function getAllInvoiceTypes() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAllInvoiceTypes = database.getAll("invoice_types")

End Function



'/*
'
'Delete Invoice Type By ID
'
'*/
Public Function deleteInvoiceType(ByVal idData As String)

    Dim database As Object
    Set database = XdbFactory.Create
    Call database.deleteById(idData, "invoice_types")

End Function
