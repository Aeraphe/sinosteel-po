Attribute VB_Name = "view_grd_vale"

'namespace=vba-files\Views\GRD



Public Function publish(grd_id As String)

    Call generete_lx_handler(grd_id)

End Function




Private Function generete_lx_handler(grd_id As String)



    Dim data As Object
    Dim grd_wkb As Workbook
    Dim new_file_name As String
    Dim grd_full_number  As String
    Dim grd_sequence As String
    Dim grd_date As String
    Dim grd_code As String
    Dim project_id As String
    Dim recipient_id As String
    Dim recipient_folder_name As String

    Dim respGRD As ADODB.Recordset
    Set respGRD = db_grd.getById(grd_id)

    recipient_id = XdbFactory.getData(respGRD, "recipent_id")
    project_id = XdbFactory.getData(respGRD, "project_id")
    grd_code = XdbFactory.getData(respGRD, "code")
    grd_date = XdbFactory.getData(respGRD, "issue_date")
    grd_sequence = XdbFactory.getData(respGRD, "sequece_number")
    grd_full_number = UCase(Trim(grd_code & grd_sequence))


    Set data = CreateObject("Scripting.Dictionary")
    Set data = get_grd_file_info(grd_id, project_id)


    Set grd_wkb = copy_data_to_grd_sheet(data, grd_full_number)
    
    'Prepar GRD Folder path
    recipient_folder_name = XdbFactory.getData(respGRD, "folder_name")
    
    Dim descktopFolderPath As String
    descktopFolderPath = h_text_file.getFolderPath("GRD_" & recipient_folder_name & "__" & format(Now, "dd_mm_yyyy_hh_mm"))
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
        
    If Not fso.FolderExists(descktopFolderPath) Then
        fso.CreateFolder descktopFolderPath
    End If
    

    
 
    
    save_file_full_path = descktopFolderPath & "\" & grd_full_number & ".xlsb"
 
    grd_wkb.SaveAs fileName:=save_file_full_path

    Dim grd_sheet As Worksheet
    Set grd_sheet = grd_wkb.Sheets("index")
    grd_sheet.name = grd_full_number
    grd_wkb.Save


End Function


'Obsolet
Private Function get_full_file_path_handler(ByVal project_id As String, ByVal recipient_folder_name As String, grd_number As String) As String

    Dim app_folders As Object
    Dim fso As Object
    Dim file_name As String
    Dim recipient_folder_path As String
    Dim folder_path As String

    Dim recipient_folder_sent_path  As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set app_folders = CreateObject("Scripting.Dictionary")
    Set app_folders = helper_app.get_projec_folders(project_id)


    recipient_folder_path = app_folders("ENG_GRD_FULL_PATH") & "\" & recipient_folder_name

    If Not fso.FolderExists(recipient_folder_path) Then
        fso.CreateFolder recipient_folder_path
    End If
    recipient_folder_sent_path = recipient_folder_path & "\" & app_folders("ENG_GRD_SENT")

    If Not fso.FolderExists(recipient_folder_sent_path) Then
        fso.CreateFolder recipient_folder_sent_path
    End If

    folder_path = recipient_folder_sent_path & "\" & grd_number & "__" & Day(Date) & "_" & Month(Date) & "_" & Year(Date)

    If Not fso.FolderExists(folder_path) Then
        fso.CreateFolder folder_path
    End If

    get_full_file_path_handler = folder_path & "\" & grd_number & ".xlsb"


End Function


Private Function copy_data_to_grd_sheet(data As Variant, ByVal grd_full_number As String) As Workbook


    Dim grd_wkb As Workbook
    Dim grd_sheet As Worksheet
    Dim file_full_path As String
    Dim doc_id As String


    Application.ScreenUpdating = False

    file_full_path = data("FULL_PATH")
    Set grd_wkb = Workbooks.Open(file_full_path)

    ' grd_wkb.visible = False

    Dim grd_tb As ListObject

    Set grd_sheet = grd_wkb.Sheets("index")

    'grd_sheet.Range("grd_tb").ClearContents

    Set grd_tb = grd_sheet.ListObjects("grd_tb")

    Set data_grd = CreateObject("Scripting.Dictionary")
    data_grd("ID") = data("GRD_ID")

    Dim respQuery As ADODB.Recordset
    Set respQuery = db_grd.getGRDItems(data_grd)

    tb_row = 1

    Do Until respQuery.EOF

        doc_id = XdbFactory.getData(respQuery, "id")

        doc_number = UCase(RemoveLineBreak(XdbFactory.getData(respQuery, "doc_number")))
        doc_extension = LCase(XdbFactory.getData(respQuery, "doc_extension"))

        grd_tb.ListColumns("Filename").DataBodyRange(tb_row).Value = doc_number & "." & doc_extension
        grd_tb.ListColumns("Name").DataBodyRange(tb_row).Value = doc_number
        grd_tb.ListColumns("Título").DataBodyRange(tb_row).Value = UCase(RemoveLineBreak(XdbFactory.getData(respQuery, "name") & " - " & XdbFactory.getData(respQuery, "description")))
        grd_tb.ListColumns("Número da Contratada").DataBodyRange(tb_row).Value = UCase(RemoveLineBreak(XdbFactory.getData(respQuery, "sinosteel_doc_number")))

        grd_tb.ListColumns("Revisão").DataBodyRange(tb_row).Value = UCase(RemoveLineBreak(XdbFactory.getData(respQuery, "rev_code")))
        grd_tb.ListColumns("Número de Páginas/Folhas").DataBodyRange(tb_row).Value = RemoveLineBreak(XdbFactory.getData(respQuery, "pages"))
        grd_tb.ListColumns("Tipo de Emissão").DataBodyRange(tb_row).Value = LCase(RemoveLineBreak(XdbFactory.getData(respQuery, "issue")))
        grd_tb.ListColumns("Formato do Papel").DataBodyRange(tb_row).Value = LCase(RemoveLineBreak(XdbFactory.getData(respQuery, "doc_format")))
        grd_tb.ListColumns("Tipo de Documento").DataBodyRange(tb_row).Value = LCase(RemoveLineBreak(XdbFactory.getData(respQuery, "doc_type_code")))

        grd_tb.ListColumns("Número GR Contratada").DataBodyRange(tb_row).Value = grd_full_number
        grd_tb.ListColumns("Primeira Emissão").DataBodyRange(tb_row).Value = get_first_doc_review_date_handler(doc_id)
        grd_tb.ListColumns("Data Realizada").DataBodyRange(tb_row).Value = Month(Date) & "/" & Day(Date) & "/" & Year(Date)



        tb_row = tb_row + 1
        respQuery.MoveNext
    Loop




    Application.ScreenUpdating = True
    Set copy_data_to_grd_sheet = grd_wkb

End Function


Private Function get_first_doc_review_date_handler(doc_id As String) As String

    Dim respQuery As ADODB.Recordset
    Set respQuery = db_documents.get_first_review(doc_id)
    first_review_date = XdbFactory.getData(respQuery, "grd_date")
    date_splited = Split(first_review_date, "-")
    get_first_doc_review_date_handler = date_splited(2) & "/" & date_splited(1) & "/" & date_splited(0)


End Function



Private Function get_grd_file_info(grd_id As String, project_id As String) As Object


    Dim respProjectGRDFile As ADODB.Recordset
    Set respProjectGRDFile = db_porject_files.get_by_type(project_id, "GRD")

    Dim path_to_forms As String
    Dim default_grd_file_name As String

    default_grd_file_name = XdbFactory.getData(respProjectGRDFile, "file_name")

    path_to_forms = config_sheet.Range("CONF_DEFAULT_FORM_PATH").Value


    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("PATH") = path_to_forms
    data("GRD_FILE_NAME") = XdbFactory.getData(respProjectGRDFile, "file_name")
    data("FULL_PATH") = path_to_forms & "\" & default_grd_file_name
    data("GRD_ID") = grd_id


    Set get_grd_file_info = data

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
