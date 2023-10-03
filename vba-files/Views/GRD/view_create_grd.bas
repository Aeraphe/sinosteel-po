Attribute VB_Name = "view_create_grd"

'namespace=vba-files\Views\GRD


Const GRD_ITEMS_TABLE = "TB_GRD_DOCS"

Private PROJECT_FOLDER As String
Private CDOC_FOLDER As String
Private GRD_FOLDER As String

Private app As Excel.Application
Private book As Excel.Workbook
Private sheet As Worksheet


Private Function load_import_excel_app(full_file_path As String)

    Set app = New Excel.Application
    app.Visible = False
    Set book = app.Workbooks.Add(full_file_path)
    Set sheet = book.Sheets("INDEX")

End Function


'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public Sub publish(ByVal grd_id As String)


    Dim grd_file_path As String
    Dim data As Object
    Dim respGRD As ADODB.Recordset
    Dim project_id As String
    Dim recipient_folder_name As String
    Dim grdDict As Object
    Set grdDict = CreateObject("Scripting.Dictionary")


    grd_file_path = config_sheet.Range("CONF_DEFAULT_FORM_PATH").Value & "\grd_padrao_sinosteel.xlsb"
    Call load_import_excel_app(grd_file_path)

    Set data = CreateObject("Scripting.Dictionary")
    data("name") = grd_id

    Set respGRD = db_grd.getById(grd_id)

    grdDict.Add "r_company", UCase(XdbFactory.getData(respGRD, "name"))
    grdDict.Add "r_person", UCase(XdbFactory.getData(respGRD, "person"))
    grdDict.Add "r_email", UCase(XdbFactory.getData(respGRD, "email"))
    grdDict.Add "r_cnpj", UCase(XdbFactory.getData(respGRD, "sup_code"))
    grdDict.Add "GRD_NUMBER", XdbFactory.getData(respGRD, "code") & XdbFactory.getData(respGRD, "sequece_number")
    grdDict.Add "GRD_DATE", XdbFactory.getData(respGRD, "issue_date")
    grdDict.Add "recipient_folder_name", XdbFactory.getData(respGRD, "folder_name")
    grdDict.Add "project_id", XdbFactory.getData(respGRD, "project_id")

    Call set_grd_header_handler(grdDict)
    Call set_grd_items_handler(data)


    'Prepar GRD Folder path
    If (grdDict("recipient_folder_name") <> "") Then
    recipient_folder_name = grdDict("recipient_folder_name") & "\"
    Else
    recipient_folder_name = "GRD"
    End If
    

    Dim grdFolderPath As String
    grdFolderPath = h_text_file.getFolderPath(recipient_folder_name)


    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(grdFolderPath) Then
        fso.CreateFolder grdFolderPath
    End If

    Dim save_file_full_path As String
    save_file_full_path = grdFolderPath & "_" & grdDict("GRD_NUMBER") & "_.xlsb"

    Call save_file_handler(save_file_full_path)
    Call Alert.Show("GRD CRIADA COM SUCESSO!!!", "", 2500)


End Sub

Private Function set_grd_header_handler(ByVal grdDict As Object)


    sheet.Range("GRD_NUMBER").Value = grdDict("GRD_NUMBER")
    sheet.Range("GRD_DATE").Value = grdDict("GRD_DATE")
    sheet.Range("GRD_RECEIVER").Value = "EMPRESA: " & grdDict("r_company") & "          CNPJ: " & grdDict("r_cnpj") & "                 RESPONSÁVEL: " & grdDict("r_person") & "            E-MAIL: " & grdDict("r_email")

    ' sheet.Range("GRD_SENDER").Value = "EMPRESA: SINOSTEEL     RESPONSÁVEL: " & Auth.user_name & "            E-MAIL: " & Auth.user_email
    sheet.Range("GRD_USER_SENDER").Value = auth.user_name & vbNewLine & "( " & auth.user_email & " )"


End Function


Private Function save_file_handler(ByVal save_file_full_path As String)
    Application.DisplayAlerts = False

    Call file_helper.DeleteFile(save_file_full_path)

    book.SaveAs save_file_full_path, FileFormat:=xlExcel12
    book.Close
    app.Quit

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Function



'obsolet
Private Function get_full_file_path_handler(ByVal project_id As String, ByVal recipient_folder_name As String) As String

    Dim app_folders As Object
    Dim fso As Object
    Dim file_name As String
    Dim recipient_folder_path As String
    Dim folder_path As String

    Dim recipient_folder_sent_path  As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set app_folders = CreateObject("Scripting.Dictionary")
    Set app_folders = helper_app.get_projec_folders(project_id)

    file_name = sheet.Range("GRD_NUMBER").Value

    recipient_folder_path = app_folders("ENG_GRD_FULL_PATH") & "\" & recipient_folder_name

    If Not fso.FolderExists(recipient_folder_path) Then
        fso.CreateFolder recipient_folder_path
    End If
    recipient_folder_sent_path = recipient_folder_path & "\" & app_folders("ENG_GRD_SENT")

    If Not fso.FolderExists(recipient_folder_sent_path) Then
        fso.CreateFolder recipient_folder_sent_path
    End If

    folder_path = recipient_folder_sent_path & "\" & file_name & "_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date)

    If Not fso.FolderExists(folder_path) Then
        fso.CreateFolder folder_path
    End If

    get_full_file_path_handler = folder_path & "\" & file_name & ".xlsx"


End Function








Private Function set_grd_items_handler(data As Object)



    sheet.Range("TB_GRD_DOCS").ClearContents

    Dim iObject As ListObject
    Dim iNewRow As ListRow

    Set iObject = sheet.ListObjects("TB_GRD_DOCS")

    Dim total_grd_rows As Long
    total_grd_rows = sheet.Range("TB_GRD_DOCS").Rows.count

    Dim respQuery As ADODB.Recordset
    Set respQuery = db_grd.getGRDItems(data)

    Dim total_items As Long


    total_items = count_records(respQuery)


    Call set_grd_table_number_of_rows(total_items)


    Dim tb_row As Long
    tb_row = 1


    Do Until respQuery.EOF

        iObject.ListColumns("ITEM").DataBodyRange(tb_row).Value = tb_row
        iObject.ListColumns("DOCUMENTO").DataBodyRange(tb_row).Value = RemoveLineBreak(XdbFactory.getData(respQuery, "doc_number"))
        iObject.ListColumns("DOCUMENTO SINOSTEEL").DataBodyRange(tb_row).Value = RemoveLineBreak(XdbFactory.getData(respQuery, "sinosteel_doc_number"))
        iObject.ListColumns("TÍTULO").DataBodyRange(tb_row).Value = RemoveLineBreak(XdbFactory.getData(respQuery, "name")) & " - " & RemoveLineBreak(XdbFactory.getData(respQuery, "description"))
        iObject.ListColumns("REV.").DataBodyRange(tb_row).Value = XdbFactory.getData(respQuery, "rev_code")
        iObject.ListColumns("TE").DataBodyRange(tb_row).Value = XdbFactory.getData(respQuery, "issue")
        iObject.ListColumns("PAGINAS").DataBodyRange(tb_row).Value = XdbFactory.getData(respQuery, "pages")
        iObject.ListColumns("MIDA").DataBodyRange(tb_row).Value = XdbFactory.getData(respQuery, "doc_media_type")
        iObject.ListColumns("TIPO").DataBodyRange(tb_row).Value = "N/A"

        tb_row = tb_row + 1
        respQuery.MoveNext
    Loop



End Function



Private Function count_records(respQuery As ADODB.Recordset) As Long

    Dim records As Long
    records = 1


    Do Until respQuery.EOF

        records = records + 1
        respQuery.MoveNext
    Loop

    respQuery.MoveFirst

    count_records = records
End Function

'/*
'
'
'This method will set the number of rows for the quantity of grd items
'
'
'*/
Private Function set_grd_table_number_of_rows(total_items As Long)

    Dim total_rows As Long
    total_rows = sheet.Range(GRD_ITEMS_TABLE).Rows.count

    Dim iObject As ListObject
    Dim iNewRow As ListRow

    Set iObject = sheet.ListObjects(GRD_ITEMS_TABLE)


    Dim i As Integer


    If (total_items > total_rows) Then

        For i = 1 To total_items - total_rows
            iObject.ListRows.Add
        Next i

    ElseIf (total_rows > total_items + 25) Then

        iObject.ListRows.item(1).delete
        Call set_grd_table_number_of_rows(total_items)
    End If

End Function





Private Function RemoveLineBreak(myString) As String
    For i = 1 To 7
        If Len(myString) > 0 Then
            If Right$(myString, 2) = vbCrLf Or Right$(myString, 2) = vbNewLine Then
                myString = Left$(myString, Len(myString) - 2)
            End If
        End If
    Next i
    RemoveLineBreak = Trim(myString)
End Function
