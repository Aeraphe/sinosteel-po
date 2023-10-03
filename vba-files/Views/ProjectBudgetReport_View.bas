Attribute VB_Name = "ProjectBudgetReport_View"

'namespace=vba-files\Views

Const FIRST_LINE = 20

'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public Sub publish(ByRef params As ADODB.Recordset)


  Dim selectedLine As Long
  Dim budget As Long
  Dim budget_rest As Long
  Dim project_id  As Long
  Dim project_name As String


  selectedLine = FIRST_LINE
  Project_Budget_Report_Sheet.Range("PROJECT_BUDGET_ITEMS_TABLE").ClearContents
  Do Until params.EOF

    project_id = XdbFactory.getData(params, "project_id")
    project_name = XdbFactory.getData(params, "project_name")

    budget_rest = XdbFactory.getData(params, "budget_rest")
    budget = XdbFactory.getData(params, "budget")
    budget_id = XdbFactory.getData(params, "budget_item_id")
    budget_name = XdbFactory.getData(params, "name")

    Project_Budget_Report_Sheet.Cells(selectedLine, "A").Value = budget_id
    Project_Budget_Report_Sheet.Cells(selectedLine, "B").Value = budget_name
    Project_Budget_Report_Sheet.Cells(selectedLine, "C").Value = XdbFactory.getData(params, "category")
    Project_Budget_Report_Sheet.Cells(selectedLine, "D").Value = XdbFactory.getData(params, "po_total")
    Project_Budget_Report_Sheet.Cells(selectedLine, "E").Value = budget
    Project_Budget_Report_Sheet.Cells(selectedLine, "F").Value = budget_rest
    Project_Budget_Report_Sheet.Cells(selectedLine, "G").Value = budget_rest / budget
    Project_Budget_Report_Sheet.Cells(selectedLine, "H").Value = XdbFactory.getData(params, "manpower")

    Project_Budget_Report_Sheet.Cells(selectedLine, "J").Value = budget_id & " - " & Left(budget_name, 7)


    selectedLine = selectedLine + 1
    params.MoveNext
  Loop

  Project_Budget_Report_Sheet.Range("PROJECT_BUDGET_REPORT_UPDATE").Value = Now()
  config_sheet.Range("CONFIG_SELECTED_PROJECT_ID").Value = project_id
  config_sheet.Range("CONFIG_SELECTED_PROJECT_NAME").Value = project_name

 ' dynamics_sheets.PivotTables("budget_category_tb").PivotCache.Refresh
 ' dynamics_sheets.PivotTables("project_budget_item_tb").PivotCache.Refresh
 ' dynamics_sheets.PivotTables("total_budget_tb").PivotCache.Refresh

'  With dynamics_sheets.PivotTables("project_budget_item_tb").PivotFields("Small String")
'    .PivotItems("(blank)").Visible = False
'  End With

'  With dynamics_sheets.PivotTables("budget_category_tb").PivotFields("Small String")
'    .PivotItems("(blank)").Visible = False
'  End With
'  With dynamics_sheets.PivotTables("total_budget_tb").PivotFields("Small String")
'    .PivotItems("(blank)").Visible = False
'  End With


End Sub

