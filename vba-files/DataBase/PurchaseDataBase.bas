Attribute VB_Name = "PurchaseDataBase"


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

    Create = database.Insert("purchase_order", data)

End Function


'/*
'
'Create Sppliers
'
'@param <Array>  data
'
'*/
Public Function InsertPurchaseItems(ByRef data As Variant) As Long

    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    InsertPurchaseItems = database.Insert("purchase_items", data)

End Function


Public Function getAll(Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "name") As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("purchase_order", searchWord, filterType)

End Function



Public Function getAllProjectPurchase(project_id As String, Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "name") As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    If (searchWord = "") Then
        sqlStrQuery = "SELECT  * FROM  purchase_order  WHERE project_id = " & project_id & " ORDER BY id DESC "
    Else

        sqlStrQuery = "SELECT  * FROM  purchase_order    Where  WHERE project_id = " & project_id & "  AND " & filterType & "   LIKE '%" & searchWord & "%'  ORDER BY id DESC "
    End If

    Set getAllProjectPurchase = database.cn.Execute(sqlStrQuery)


End Function



Public Function getPurchaseFromBudget(project_id As String, budget_item_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT  * FROM  purchase_order  WHERE project_id = " & project_id & " AND budget_item_id = " & budget_item_id & " ORDER BY id DESC "

    Set getPurchaseFromBudget = database.cn.Execute(sqlStrQuery)


End Function
