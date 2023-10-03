Attribute VB_Name = "view_ld"


'namespace=vba-files\Views\Reports



Const FIRST_DATA_ROW = 7
Const HEADER_ROW = 6

Private import_app As Excel.Application
Private import_book As Excel.Workbook
Private import_sheet As Worksheet
Private import_file_info As Object



Function publish(ByVal project_selected_id As String, ByVal ld_excel_file_path As String)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Load UserFormAlert
    UserFormAlert.Label1.Caption = "Gerando LD"
    UserFormAlert.Show

    Call load_import_excel_app(ld_excel_file_path, "INDEX")
    Call populate_excel_file(project_selected_id)
    Call save_report_file_handler(project_selected_id)


    UserFormAlert.Label1.Caption = "LISTA DE DOCUMENTOS (LD)"
    UserFormAlert.labelInfo.Caption = "GERADA COM SUCESSO!!"
    Call Xhelper.waitMs(2000)

    Unload UserFormAlert

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic




End Function

Private Function load_import_excel_app(ByVal full_file_path As String, ByVal sheet_name As String)

    Set import_app = New Excel.Application
    import_app.Visible = False
    Set import_book = import_app.Workbooks.Add(full_file_path)
    On Error GoTo ErrorHandler
    Set import_sheet = import_book.Sheets(sheet_name)
    Exit Function

ErrorHandler:
    Call Alert.Show("Arquivo Selecionado não Contem a Planilha INDEX requerida", "", 3000)
    import_app.Quit
    End
End Function


Private Function populate_excel_file(project_id As String)


    import_app.ScreenUpdating = False
    import_app.Calculation = xlCalculationManual

    import_sheet.Range("CREATE_ON").Value = Now
    import_sheet.Range("CREATE_BY").Value = auth.user_name & vbNewLine & auth.user_email

    Call create_ld_header
    Call create_ld_rows(project_id)

    import_app.ScreenUpdating = True
    import_app.Calculation = xlCalculationAutomatic


End Function



Private Function create_ld_header()



    import_sheet.Cells(HEADER_ROW, 1) = "ITEM"
    import_sheet.Cells(HEADER_ROW, 2) = "N FORNECEDOR" ' get_db_value(respQuery, "doc_number")
    import_sheet.Cells(HEADER_ROW, 3) = "N SINOSTEEL" 'get_db_value(respQuery, "sinosteel_doc_number")
    import_sheet.Cells(HEADER_ROW, 4) = "TITULO PRIMARIO" ' get_db_value(respQuery, "name")
    import_sheet.Cells(HEADER_ROW, 5) = "TITULO ECUNDARIO" 'get_db_value(respQuery, "description")
    import_sheet.Cells(HEADER_ROW, 6) = "CODIGO DOC" 'get_db_value(respQuery, "doc_type_code")
    import_sheet.Cells(HEADER_ROW, 7) = "ITEM CONTRATO" 'get_db_value(respQuery, "contract_item")
    import_sheet.Cells(HEADER_ROW, 8) = "FORMATO" 'get_db_value(respQuery, "doc_format")
    import_sheet.Cells(HEADER_ROW, 9) = "PAGINAS" 'get_db_value(respQuery, "pages")
    import_sheet.Cells(HEADER_ROW, 10) = "PRIMEIRA REV" 'first_rev
    import_sheet.Cells(HEADER_ROW, 11) = "PRIMERA TE" 'first_issue
    import_sheet.Cells(HEADER_ROW, 12) = "GRD PRIMEIRA REV" 'get_db_value(respQuery, "first_review_grd")
    import_sheet.Cells(HEADER_ROW, 13) = "DATA GRD PRIMEIRA REV" 'DateHelpers.FormatSQliteToDate(get_db_value(respQuery, "first_review_grd_date"))
    import_sheet.Cells(HEADER_ROW, 14) = "REV ATUAL" 'actual_rev
    import_sheet.Cells(HEADER_ROW, 15) = "TE ATUAL" 'acutal_issue
    import_sheet.Cells(HEADER_ROW, 16) = "GRD REV ATUAL" 'get_db_value(respQuery, "last_review_grd")
    import_sheet.Cells(HEADER_ROW, 17) = "DATA GRD REV ATUAL" 'get_db_value(respQuery, "last_review_grd_date")
    import_sheet.Cells(HEADER_ROW, 18) = "DISCIPLINA" 'get_db_value(respQuery, "discipline")
    import_sheet.Cells(HEADER_ROW, 19) = "DISCIPLINA_CODE" ' get_db_value(respQuery, "discipline_code")
    import_sheet.Cells(HEADER_ROW, 20) = "CATEGORIA" 'get_db_value(respQuery, "category")
    import_sheet.Cells(HEADER_ROW, 21) = "PASTA" 'get_db_value(respQuery, "folder")
    import_sheet.Cells(HEADER_ROW, 22) = "STATUS" 'get_db_value(respQuery, "last_review_status")
    import_sheet.Cells(HEADER_ROW, 23) = "STATUS DATA" 'DateHelpers.FormatSQliteToDate(get_db_value(respQuery, "last_review_status_date"))
    import_sheet.Cells(HEADER_ROW, 24) = "OBS REV ATUAL" 'acutal_issue
    import_sheet.Cells(HEADER_ROW, 25) = "OBS REV ATUAL" 'get_db_value(respQuery, "last_review_obs")
    import_sheet.Cells(HEADER_ROW, 26) = "GRD RECEBIDA REV ATUAL" 'get_db_value(respQuery, "last_review_grd_receive")
    import_sheet.Cells(HEADER_ROW, 27) = "GRD RECEBIDA DATA REV ATUAL" 'DateHelpers.FormatSQliteToDate(get_db_value(respQuery, "last_review_grd_date_receive"))
    import_sheet.Cells(HEADER_ROW, 28) = "REV ID" 'get_db_value(respQuery, "last_review_id")





End Function


Private Function create_ld_rows(ByVal project_id As String)


    Dim first_rev As String
    Dim first_issue As String
    Dim acutal_issue As String
    Dim actual_rev As String
    Dim doc_number As String
    Dim respQuery As ADODB.Recordset
    Dim row As Long
    Dim items_count As Long

    import_sheet.Range("LD_SINOSTEEL_TB").ClearContents

    Set respQuery = db_ld_report.generate(project_id)
 

    items_count = 1
    row = FIRST_DATA_ROW

    Do Until respQuery.EOF

        first_rev = get_db_value(respQuery, "first_review")
        first_issue = get_db_value(respQuery, "first_issue")
        actual_rev = get_db_value(respQuery, "last_review")
        acutal_issue = get_db_value(respQuery, "last_issue")
        doc_number = get_db_value(respQuery, "doc_number")

        import_sheet.Cells(row, 1) = items_count
        import_sheet.Cells(row, 2) = doc_number
        import_sheet.Cells(row, 3) = get_db_value(respQuery, "sinosteel_doc_number")
        import_sheet.Cells(row, 4) = get_db_value(respQuery, "name")
        import_sheet.Cells(row, 5) = get_db_value(respQuery, "description")
        import_sheet.Cells(row, 6) = get_db_value(respQuery, "doc_type_code")
        import_sheet.Cells(row, 7) = get_db_value(respQuery, "contract_item")
        import_sheet.Cells(row, 8) = get_db_value(respQuery, "doc_format")
        import_sheet.Cells(row, 9) = get_db_value(respQuery, "pages")
        import_sheet.Cells(row, 10) = first_rev
        import_sheet.Cells(row, 11) = first_issue
        import_sheet.Cells(row, 12) = get_db_value(respQuery, "first_review_grd")
        import_sheet.Cells(row, 13) = DateHelpers.FormatSQliteToDate(get_db_value(respQuery, "first_review_grd_date"))
        import_sheet.Cells(row, 14) = actual_rev
        import_sheet.Cells(row, 15) = acutal_issue
        import_sheet.Cells(row, 16) = get_db_value(respQuery, "last_review_grd")
        import_sheet.Cells(row, 17) = get_db_value(respQuery, "last_review_grd_date")
        import_sheet.Cells(row, 18) = get_db_value(respQuery, "discipline")
        import_sheet.Cells(row, 19) = get_db_value(respQuery, "discipline_code")
        import_sheet.Cells(row, 20) = get_db_value(respQuery, "category")
        import_sheet.Cells(row, 21) = get_db_value(respQuery, "folder")
        import_sheet.Cells(row, 22) = get_db_value(respQuery, "last_review_status")
        import_sheet.Cells(row, 23) = DateHelpers.FormatSQliteToDate(get_db_value(respQuery, "last_review_status_date"))
        import_sheet.Cells(row, 24) = acutal_issue
        import_sheet.Cells(row, 25) = get_db_value(respQuery, "last_review_obs")
        import_sheet.Cells(row, 26) = get_db_value(respQuery, "last_review_grd_receive")
        import_sheet.Cells(row, 27) = XdbFactory.getData(respQuery, "last_review_grd_date_receive", "SQL_DATE")
        import_sheet.Cells(row, 28) = get_db_value(respQuery, "last_review_id")

        UserFormAlert.labelInfo.Caption = "Incluindo Documento: " & doc_number & " de [ " & items_count & " ]"
        UserFormAlert.Repaint
        respQuery.MoveNext

        
        items_count = items_count + 1
        row = row + 1
    Loop



End Function


'/*
'
'
'
'
'
'*/
Private Function get_db_value(respQuery As ADODB.Recordset, ByVal prop As String) As String
    Dim Value As Variant

    Value = respQuery.fields.item(prop)

    get_db_value = Xhelper.iff(IsNull(Value), "", Value)

End Function

Private Function save_report_file_handler(project_id As String)

    Dim app_folders As Object
    Dim fso As Object
    Dim file_name As String
    Dim save_file_full_path As String

    Application.DisplayAlerts = False


    file_name = "LD__" & format(Now(), "yyyy_MM_dd_hh_mm_ss") & ".xlsx"


    Set fso = CreateObject("Scripting.FileSystemObject")

    Set app_folders = CreateObject("Scripting.Dictionary")
    Set app_folders = helper_app.get_projec_folders(project_id)

    If Not fso.FolderExists(app_folders("ENG_REPORTS_FULL_PATH")) Then
        fso.CreateFolder app_folders("ENG_REPORTS_FULL_PATH")
    End If

    save_file_full_path = app_folders("ENG_REPORTS_FULL_PATH") & "\" & file_name

    Debug.Print save_file_full_path
    import_book.SaveAs save_file_full_path
    import_book.Close
    import_app.Quit

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Function






