Attribute VB_Name = "view_doc_flow_report"



'namespace=vba-files\Views\Reports

Public project_selected_id
Private import_app As Excel.Application
Private import_book As Excel.Workbook
Private import_sheet As Worksheet
Private import_file_info As Object



Function publish(project_selected_id As String, date_selected As String)

    Application.ScreenUpdating = False
    Dim excel_report_full_file_path As String
    Dim save_file_folder As String
    Dim save_file_full_path As String


    excel_report_full_file_path = config_sheet.Range("CONF_DEFAULT_FORM_PATH").Value & "\RELATORIO_PADRAO_DOCUMENTOS_COMENTADOS.xlsb"

    Call load_import_excel_app(excel_report_full_file_path)

    Call populate_report_file(project_selected_id, date_selected)


    Call save_report_file_handler(project_selected_id)

End Function


Private Function save_report_file_handler(project_id As String)

    Dim app_folders As Object
    Dim fso As Object
    Dim file_name As String
    Dim save_file_full_path As String

    Application.DisplayAlerts = False


    file_name = "RELATORIO_DOCS_RECEBIDOS_COMENTADOS_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date) & ".xlsx"


    Set fso = CreateObject("Scripting.FileSystemObject")

    Set app_folders = CreateObject("Scripting.Dictionary")
    Set app_folders = helper_app.get_projec_folders(project_id)

    If Not fso.FolderExists(app_folders("ENG_REPORTS_FULL_PATH")) Then
        fso.CreateFolder app_folders("ENG_REPORTS_FULL_PATH")
    End If

    save_file_full_path = app_folders("ENG_REPORTS_FULL_PATH") & "\" & file_name

    Call file_helper.DeleteFile(save_file_full_path)

    import_book.SaveAs save_file_full_path
    import_book.Close
    import_app.Quit

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Function

Private Function load_import_excel_app(full_file_path As String)

    Set import_app = New Excel.Application
    import_app.Visible = False
    Set import_book = import_app.Workbooks.Add(full_file_path)
    Set import_sheet = import_book.Sheets("index")

End Function




Private Function populate_report_file(project_id As String, date_selected As String)


    Dim total_rows As Long
    Dim i  As Long
    Dim prop_name As String
    Dim prop_value  As Variant
    Dim doc_id As String
    Dim iObject As ListObject
    Dim iNewRow As ListRow
    Dim data As Object
    Dim doc_number As String
    Dim status_date As String
    Dim respQuery As ADODB.Recordset
    Dim next_review As Variant
    Dim next_issue As String
    Dim status As String
    Dim old_rev As Variant
    Dim old_issue As String
    Dim is_certificated As String


    Load UserFormAlert
    UserFormAlert.Label1.Caption = "Gerando Relatório"
    UserFormAlert.Show


    total_rows = import_sheet.Range("TD_DOCS").Rows.count
    Set iObject = import_sheet.ListObjects("TD_DOCS")

    import_sheet.Range("CREATE_ON").Value = Now
    import_sheet.Range("CREATE_BY").Value = auth.user_name & vbNewLine & auth.user_email

    Set data = CreateObject("Scripting.Dictionary")
    data("name") = project_id
    data("selected_date") = DateHelpers.FormatDateToSQlite(date_selected)

    Set respQuery = db_document_reports.get_documents_today_status_change(data)

    i = 0

    Do Until respQuery.EOF

        i = i + 1

        status_date = DateHelpers.FormatSQliteToDate(XdbFactory.getData(respQuery, "status_date"))
        next_review = XdbFactory.getData(respQuery, "next_review")
        next_issue = XdbFactory.getData(respQuery, "next_issue")
        status = XdbFactory.getData(respQuery, "status")
        old_rev = XdbFactory.getData(respQuery, "rev_code")
        old_issue = XdbFactory.getData(respQuery, "issue")
        
        On Error Resume Next
        next_review = CInt(next_review)
        
        On Error Resume Next
        old_rev = CInt(old_rev)
        
        iObject.ListColumns("ITEM").DataBodyRange(i).Value = i
        iObject.ListColumns("DOCUMENTO").DataBodyRange(i).Value = XdbFactory.getData(respQuery, "doc_number")
        iObject.ListColumns("TÍTULO").DataBodyRange(i).Value = XdbFactory.getData(respQuery, "name") & " - " & XdbFactory.getData(respQuery, "description")
        iObject.ListColumns("TIPO").DataBodyRange(i).Value = XdbFactory.getData(respQuery, "contract_item") & " - " & XdbFactory.getData(respQuery, "category")
        iObject.ListColumns("FORNECEDOR").DataBodyRange(i).Value = XdbFactory.getData(respQuery, "supplier")
        iObject.ListColumns("OBSERVAÇÃO").DataBodyRange(i).Value = XdbFactory.getData(respQuery, "obs")
        iObject.ListColumns("REV.").DataBodyRange(i).Value = old_rev
        iObject.ListColumns("TE").DataBodyRange(i).Value = old_issue
        iObject.ListColumns("REVISÃO").DataBodyRange(i).Value = next_review
        iObject.ListColumns("EMISSÃO").DataBodyRange(i).Value = next_issue
        iObject.ListColumns("STATUS").DataBodyRange(i).Value = status
        iObject.ListColumns("DATA").DataBodyRange(i).Value = DateAdd("d", 7, status_date)
        If (next_review = old_rev And old_issue = next_issue And status = "APR") Then
          is_certificated = "CERTIFICADO"
        
        ElseIf (VarType(old_rev) = vbString And VarType(next_review) = vbInteger And status <> "APR") Then
         is_certificated = "EMITIR CERTIFICADO"
         Else
             is_certificated = "EMITIR"
        End If
        
        iObject.ListColumns("CERTIFICADO").DataBodyRange(i).Value = is_certificated



        UserFormAlert.Label1.Caption = "Gerando Relatório : " & i
        UserFormAlert.Repaint

        respQuery.MoveNext

    Loop


    Unload UserFormAlert

End Function


