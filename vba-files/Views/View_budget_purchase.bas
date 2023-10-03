Attribute VB_Name = "View_budget_purchase"

'namespace=vba-files\Views

Const FIRST_LINE = 7

'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public Sub publish(ByRef params As ADODB.Recordset)


    Dim selectedLine As Long
    Dim purchase_id As Long
    Dim budget_rest As Long
    Dim project_id  As Long
    Dim project_name As String


    selectedLine = FIRST_LINE

    budget_purchase.Range("PROJECT_BUDGET_ITEMS_TABLE3").ClearContents
    
    Do Until params.EOF

        purchase_id = XdbFactory.getData(params, "id")


        budget_purchase.Cells(selectedLine, "A").Value = purchase_id
        budget_purchase.Cells(selectedLine, "B").Value = XdbFactory.getData(params, "po_code")
        budget_purchase.Cells(selectedLine, "C").Value = XdbFactory.getData(params, "description")
        budget_purchase.Cells(selectedLine, "D").Value = XdbFactory.getData(params, "doc_issuance_date")
        budget_purchase.Cells(selectedLine, "E").Value = XdbFactory.getData(params, "iconterms")
        budget_purchase.Cells(selectedLine, "F").Value = XdbFactory.getData(params, "payment_date")
        budget_purchase.Cells(selectedLine, "G").Value = XdbFactory.getData(params, "currency")
        budget_purchase.Cells(selectedLine, "H").Value = XdbFactory.getData(params, "delivery_time")
        budget_purchase.Cells(selectedLine, "I").Value = XdbFactory.getData(params, "obs")

        selectedLine = selectedLine + 1
        params.MoveNext
    Loop



End Sub

