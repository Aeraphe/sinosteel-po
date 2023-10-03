Attribute VB_Name = "BudgetTypeDataBase"


'namespace=vba-files\DataBase




Public Function getAll() As Variant

    Dim database As Object
    Set database = XdbFactory.Create

    Set getAll = database.getAll("budget_types")

End Function

