Attribute VB_Name = "ProjectBudgetItemDataBase"


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
    On Error Resume Next
    Create = database.Insert("projects_budgets_items", data)

End Function



Public Function getProjectBudgetItems(project_id As String) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = "SELECT  item.id,item.name,item.category,pbi.manpower as hh, pbi.price FROM  projects_budgets_items AS pbi  INNER JOIN budgets_items as item ON pbi.budget_item_id = item.id  Where  pbi.project_id = " & project_id & "  ORDER BY pbi.project_id DESC "


    Set getProjectBudgetItems = database.cn.Execute(sqlStrQuery)


End Function



'/*
'
'Delete budget By ID
'
'*/
Public Function delete(ByRef data As Variant)

    Dim rs As ADODB.Recordset
    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    sqlStrQuery = " DELETE   FROM  projects_budgets_items  Where project_id = " & data("project_id") & " AND  budget_item_id = " & data("budget_item_id")

    'Set the recor set
    Set rs = New ADODB.Recordset
    Set rs = database.cn.Execute(sqlStrQuery)


End Function
