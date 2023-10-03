Attribute VB_Name = "Report_ProjectBudget_DataBase"


'namespace=vba-files\DataBase




Public Function GetProjectBudgets(params As Variant) As Variant


    Dim sqlStrQuery As String
    Dim database As Object

    Set database = XdbFactory.Create

    Set GetProjectBudgets = database.SelectX("report_project_budget", params)

End Function
