VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_check_review_form 
   Caption         =   "Verificar Rev."
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   OleObjectBlob   =   "doc_check_review_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_check_review_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents


Public total_file_rows


Private Sub UserForm_Activate()
   If (auth.is_logged) Then
     

   Else
      doc_add_form.Hide
 
      Call Alert.Show("Favor Logar no Sistema", "", 3000)

   End If
End Sub



Private Sub check_btn_Click()


   If (auth.is_authorized("SUPER_ADMIN")) Then
     
  doc_review_check_sheet.Activate
  doc_review_check_sheet.Cells(3, "B") = 0
  doc_review_check_sheet.Cells(4, "B") = 0
  doc_review_check_sheet.Cells(2, "O") = ""
  doc_review_check_sheet.Range("REPORT_VERSION_TB").ClearContents
  n = 2

  'Clear filter data
  On Error Resume Next
  doc_review_check_sheet.ShowAllData

  Application.ScreenUpdating = False
  Dim types As String
  Dim site As String

  Call Alert.Show("Iniciando a Busca de Arquivos", "", 2500)
  info_lb.Caption = "Iniciando a Busca de Arquivos"
  doc_check_review_form.Repaint

  While config_sheet.Cells(n, "H").Value <> ""
    types = config_sheet.Cells(n, "F").Value
    site = config_sheet.Cells(n, "G").Value

    title_lb.Caption = "Buscando nas Pastas"
    info_lb.Caption = " Pasta " & site & "  " & types
    doc_check_review_form.Repaint

    Call get_doc_handler(config_sheet.Cells(n, "H").Value, types, site)
    n = n + 1
  Wend


  Call Alert.Show("Calculando o total de Arquivos", "", 2500)
  title_lb.Caption = "Calculando o total de Arquivos"
  doc_check_review_form.Repaint

  Dim count_total_lines As Integer
  count_total_lines = 7
  While doc_review_check_sheet.Cells(count_total_lines, "A").Value <> ""
    count_total_lines = count_total_lines + 1

  Wend
  total_file_rows = count_total_lines


  Call Alert.Show("Limpando as Rev. de Letras", "", 2500)
  title_lb.Caption = "Limpando as Rev. de Letras"
  doc_check_review_form.Repaint

  clear_letter_review

  Call Alert.Show("Identificando a Mairor Rev. em Numero", "", 2500)
  title_lb.Caption = "Identificando a Mairor Rev. em Numero"
  doc_check_review_form.Repaint
  max_number_review

  Call Alert.Show("Identificando a Mairor Rev. em Letra", "", 2500)
  title_lb.Caption = "Identificando a Mairor Rev. em Letra"
  doc_check_review_form.Repaint
  max_letter_review

  Call Alert.Show("Apagando Linhas Vazias", "", 2500)
  title_lb.Caption = "Apagando Linhas Vazias"
  doc_check_review_form.Repaint
  delete_rows



  Call Alert.Show("Comparando com a LD", "", 2500)
  title_lb.Caption = "Comparando com a LD"
  doc_check_review_form.Repaint
  compare_review_form_ld


  Call Alert.Show("Verificando Documentos em Hold", "", 2500)
  title_lb.Caption = "Verificando Documentos em Hold"
  doc_check_review_form.Repaint
  Call check_document_in_hold

  Me.Hide



  Application.ScreenUpdating = True
  Call Alert.Show("Busca dos documentos Finalizada", "", 2500)
  doc_review_check_sheet.Cells(2, "O") = Date
  
  Else
   Call Alert.Show("Você não tem permissão para efetuar esta operação", "", 2500)
End If
End Sub



Private Function delete_rows()

  Dim line As Integer
  line = 7

  Dim clear_rows_count As Integer
  clear_rows_count = 1
  While line <= total_file_rows

    If (doc_review_check_sheet.Cells(line, "A").Value = "DELETE") Then

      info_lb.Caption = " Linhas apagadas " & "--> [Total: " & clear_rows_count & "]"
      doc_check_review_form.Repaint

      Rows(line & ":" & line).delete
      clear_rows_count = clear_rows_count + 1
      line = line - 1
    End If

    line = line + 1
  Wend

End Function

Private Function compare_review_form_ld()

  Dim start_row_report As Integer
  start_row_report = 7

  Dim start_row_ld As Integer
  start_row_ld = 3
  Dim report_row As Integer
  report_row = start_row_report
  Dim ld_row As Integer
  Dim dr As String
  Dim dr_ld As String
  Dim dr_rev As Variant
  Dim dr_ld_rev As Variant

  Dim count_files As Integer
  count_files = 1
  While doc_review_check_sheet.Cells(report_row, "A").Value <> ""
    dr = Trim(UCase(doc_review_check_sheet.Cells(report_row, "A").Value))
    dr_rev = Trim(UCase(doc_review_check_sheet.Cells(report_row, "B").Value))

    ld_row = start_row_ld
    info_lb.Caption = dr & " --> [Total: " & count_files & "]"
    doc_check_review_form.Repaint

    not_found = True
    While lds_sheet.Cells(ld_row, "A").Value <> ""
      dr_ld = Trim(UCase(lds_sheet.Cells(ld_row, "B").Value))
      dr_ld_rev = Trim(UCase(lds_sheet.Cells(ld_row, "R").Value))
      issue = Trim(UCase(lds_sheet.Cells(ld_row, "S").Value))

      If (dr = dr_ld) Then
        not_found = False
        doc_review_check_sheet.Cells(report_row, "C").Value = dr_ld_rev 'Rev LD
        doc_review_check_sheet.Cells(report_row, "I").Value = lds_sheet.Cells(ld_row, "I").Value 'Title LD
        doc_review_check_sheet.Cells(report_row, "D").Value = issue
        doc_review_check_sheet.Cells(report_row, "J").Value = lds_sheet.Cells(ld_row, "H").Value 'Item contrato LD
        doc_review_check_sheet.Cells(report_row, "E").Value = lds_sheet.Cells(ld_row, "T").Value 'Data LD
        On Error Resume Next
        doc_review_check_sheet.Cells(report_row, "F").Value = Date - CDate(lds_sheet.Cells(ld_row, "T").Value) 'Data LD
        doc_review_check_sheet.Cells(report_row, "L").Value = lds_sheet.Cells(ld_row, "AB").Value 'SDS LD


        If (issue = "H") Then

          doc_review_check_sheet.Cells(report_row, "O").Value = "CANCELADO"
        Else

          If (dr_rev = dr_ld_rev) Then

            doc_review_check_sheet.Cells(report_row, "O").Value = "OK"
          ElseIf (dr_ld_rev = "") Then
            doc_review_check_sheet.Cells(report_row, "O").Value = "SEM REV. NA LD"
          Else
            doc_review_check_sheet.Cells(report_row, "O").Value = "REV. DIFERENTE"
          End If

        End If



      End If



      ld_row = ld_row + 1
    Wend
    If (not_found) Then
      doc_review_check_sheet.Cells(report_row, "O").Value = "NÃO ENCONTRADO NA LD"
    End If
    report_row = report_row + 1
    count_files = count_files + 1
  Wend


End Function

'/*
'
'
'
'
'*/
Private Function get_doc_handler(folder_path As String, types As String, site As String)



  Dim files_dict As Object
  Set files_dict = CreateObject("Scripting.Dictionary")
  Dim error_row As Integer
  error_row = 7
  While doc_review_check_sheet.Cells(error_row, "T").Value <> ""
    error_row = error_row + 1
  Wend


  Dim file_name_arr As Integer


  Set files_dict = file_helper.get_files_from_folders(folder_path)

  firstline = 7
  n = firstline
  If (doc_review_check_sheet.Cells(firstline, "A").Value <> "") Then

    While doc_review_check_sheet.Cells(n, "A").Value <> ""
      n = n + 1
    Wend
  End If


  Dim doc_row_count As Integer

  doc_row_count = n

  For Each varKey In files_dict.Keys()
    If (varKey <> "count") Then
      file = Split(UCase(files_dict(varKey)), ".")

      If (file(1) = "DWG" Or file(1) = "XLS" Or file(1) = "DOC" Or file(1) = "DOCX" Or file(1) = "XLSX") Then
        file_name = Split(UCase(files_dict(varKey)), "_REV_")
        On Error GoTo ErrorHandler
        file_name_arr = UBound(file_name) - LBound(file_name) + 1
        If (file_name_arr = 2) Then
          extension = Split(file_name(1), ".")

          info_lb.Caption = "Incluindo: " & file_name(0) & "  De  " & n
          doc_check_review_form.Repaint

          doc_review_check_sheet.Cells(doc_row_count, "A") = file_name(0)
          doc_review_check_sheet.Cells(doc_row_count, "B") = extension(0) 'Rev
          doc_review_check_sheet.Cells(doc_row_count, "G") = extension(1) 'Extensao
          doc_review_check_sheet.Cells(doc_row_count, "H") = types
          doc_review_check_sheet.Cells(doc_row_count, "K") = site
          doc_row_count = doc_row_count + 1
        Else
          doc_review_check_sheet.Cells(error_row, "T") = file_name(0)
          error_row = error_row + 1
        End If

      Else


      End If
ErrorHandler:
      n = n + 1
    End If
  Next '

  doc_review_check_sheet.Cells(3, "B") = n + doc_review_check_sheet.Cells(3, "B")
  doc_review_check_sheet.Cells(4, "B") = files_dict("count") + doc_review_check_sheet.Cells(4, "B")




End Function


Private Function clear_letter_review()



  total_lines = total_file_rows
  Dim total_delete_doc As Integer
  total_delete_doc = 1
  n = 7

  Dim doc_rev As Variant
  While n <= total_lines
    dr = doc_review_check_sheet.Cells(n, "A").Value

    doc_rev = doc_review_check_sheet.Cells(n, "B").Value
    On Error Resume Next
    doc_rev = CInt(doc_review_check_sheet.Cells(n, "B").Value)

    If (IsNumeric(doc_rev) And dr <> "DELETE") Then
      m = 7
      While m <= total_lines

        dr_next = doc_review_check_sheet.Cells(m, "A").Value
        doc_rev_2 = doc_review_check_sheet.Cells(m, "B").Value
        On Error Resume Next
        doc_rev_2 = CInt(doc_review_check_sheet.Cells(m, "B").Value)

        If (dr = dr_next And Not IsNumeric(doc_rev_2) And dr <> "DELETE") Then
          Rows(m & ":" & m).ClearContents
          info_lb.Caption = " Apagando Revisões Obsoletas: " & doc_review_check_sheet.Cells(m, "A").Value & "--> [Total: " & total_delete_doc & "]"
          doc_check_review_form.Repaint
          total_delete_doc = total_delete_doc + 1
          doc_review_check_sheet.Cells(m, "A").Value = "DELETE"



        End If
        m = m + 1
      Wend

    End If
    n = n + 1
  Wend

End Function


Private Function max_number_review()


  Dim total_delete_doc As Integer
  total_delete_doc = 1

  total_lines = total_file_rows

  n = 7

  While n <= total_lines
    dr = Trim(UCase(doc_review_check_sheet.Cells(n, "A").Value))
    If (dr = "") Then
      Exit Function
    End If
    rev = doc_review_check_sheet.Cells(n, "B").Value
    If (IsNumeric(rev) And dr <> "DELETE") Then
      m = 7
      While m <= total_lines
        dr_next = Trim(UCase(doc_review_check_sheet.Cells(m, "A").Value))

        If (dr = dr_next And dr <> "DELETE") Then


          info_lb.Caption = " Apagando Revisões Obsoletas: " & doc_review_check_sheet.Cells(m, "A").Value & "--> [Total: " & total_delete_doc & "]"
          doc_check_review_form.Repaint
          total_delete_doc = total_delete_doc + 1


          rev_next = doc_review_check_sheet.Cells(m, "B").Value
          If (rev < rev_next) Then
            Rows(n & ":" & n).ClearContents
            doc_review_check_sheet.Cells(n, "A").Value = "DELETE"

          ElseIf (rev > rev_next) Then

            Rows(m & ":" & m).ClearContents
            doc_review_check_sheet.Cells(m, "A").Value = "DELETE"


          End If

        End If
        m = m + 1
        If (dr_next = "") Then
          m = 1 + total_lines
        End If
      Wend

    End If
    n = n + 1
  Wend

End Function


Private Function max_letter_review()


  Dim total_delete_doc As Integer
  total_delete_doc = 1

  total_lines = total_file_rows
  n = 7

  While n <= total_lines
    dr = doc_review_check_sheet.Cells(n, "A").Value
    If (dr = "") Then
      Exit Function
    End If

    If (Not IsNumeric(doc_review_check_sheet.Cells(n, "B").Value)) Then
      rev = Asc(doc_review_check_sheet.Cells(n, "B").Value)
      m = 7
      While m <= total_lines
        dr_next = doc_review_check_sheet.Cells(m, "A").Value

        If (dr = dr_next And dr <> "") Then

          info_lb.Caption = " Apagando Revisões Obsoletas: " & doc_review_check_sheet.Cells(m, "A").Value & "--> [Total: " & total_delete_doc & "]"
          doc_check_review_form.Repaint
          total_delete_doc = total_delete_doc + 1

          rev_next = Asc(doc_review_check_sheet.Cells(m, "B").Value)
          If (rev < rev_next) Then
            Rows(n & ":" & n).ClearContents
            doc_review_check_sheet.Cells(n, "A").Value = "DELETE"

          ElseIf (rev > rev_next) Then

            Rows(m & ":" & m).ClearContents
            doc_review_check_sheet.Cells(m, "A").Value = "DELETE"


          End If


        End If
        m = m + 1
        If (dr_next = "") Then
          m = 1 + total_lines
        End If
      Wend

    End If
    n = n + 1
  Wend

End Function


Private Function check_document_in_hold()


  Dim total_docs_for_check As Long
  Dim total_docs_in_hold As Long
  Dim i  As Long
  Dim j  As Long

  total_docs_in_hold = docs_in_hold_sheet.Range("DOC_HOLD_TB").Rows.count

  total_docs_for_check = doc_review_check_sheet.Range("REPORT_VERSION_TB").Rows.count

  Dim tb_check_docs As ListObject
  Set tb_check_docs = doc_review_check_sheet.ListObjects("REPORT_VERSION_TB")

  Dim tb_hold As ListObject
  Set tb_hold = docs_in_hold_sheet.ListObjects("DOC_HOLD_TB")

  For i = 1 To total_docs_for_check

    info_lb.Caption = " VERIFICANDO DOCS. EM HOLD: " & tb_check_docs.ListColumns("Desenhos").DataBodyRange(i).Value
   doc_check_review_form.Repaint

    For j = 1 To total_docs_in_hold

      If (tb_check_docs.ListColumns("Desenhos").DataBodyRange(i).Value = tb_hold.ListColumns("Desenhos").DataBodyRange(j).Value) Then

        tb_check_docs.ListColumns("Status").DataBodyRange(i).Value = tb_hold.ListColumns("STATUS").DataBodyRange(j).Value
        Exit For
      End If

    Next j


  Next i

End Function
