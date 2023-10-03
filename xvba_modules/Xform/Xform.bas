Attribute VB_Name = "Xform"

'namespace=xvba_modules\Xform





'/*
'
'Size the ListBox Columns base on content
'For use create a hidden Label named lblHidden with Autosize: true and wordwarp:false
'
'*/
Public Function SetColumnWidths(ByRef listBox As Variant, ByRef lblHidden As Variant) As String

   Dim i As Long, j As Long
   Dim sColWidths As String
   Dim dMax As Double

   If (listBox.ColumnCount) Then
      For i = 0 To listBox.ColumnCount - 1
         For j = 0 To listBox.ListCount - 1
            lblHidden.Caption = listBox.column(i, j) & "MM"
            If dMax < lblHidden.Width Then
               dMax = lblHidden.Width
            End If
         Next j
         sColWidths = sColWidths & CLng(dMax) & ";"
         dMax = 0
      Next i

      listBox.ColumnWidths = sColWidths
      lblHidden.Caption = ""
      SetColumnWidths = sColWidths
   End If


End Function


Public Function create_form_header_label(header_label As MSForms.Label, sColWidths As String, arrHeaders)

   Dim header_size As Variant
   Dim arr_size As Variant
   Dim i As Long
   Dim header_title_size As Long
   Dim space_size As Long
   Dim max_header_size As Long

   header_size = Split(sColWidths, ";")
   arr_size = UBound(arrHeaders) - LBound(arrHeaders) + 1

   Dim label_width  As Long
   Dim header As String

   Dim char_inch_size As Long
   header_label.Caption = ""
   header_label.Caption = "1"
   char_inch_size = CLng(header_label.Width)

   For i = 0 To arr_size - 1

      header_title_size = Len(arrHeaders(i)) * char_inch_size

      max_header_size = CLng(header_size(i))

      label_width = max_header_size + label_width
      If (max_header_size = header_title_size) Then

         space_size = max_header_size + char_inch_size
      ElseIf (max_header_size > header_title_size) Then
         space_size = CLng(max_header_size / char_inch_size - header_title_size / char_inch_size)
      Else
         space_size = header_title_size / char_inch_size
      End If

      If (i = 0) Then

         header = Space(4) & arrHeaders(i) & Space(space_size)
      ElseIf (i < arr_size - 1) Then '

         header = header & Space(4) & arrHeaders(i) & Space(space_size)
      Else
         header = header & arrHeaders(i)
      End If
      header_label.Caption = header
   Next i

   header_label.Width = label_width

End Function
'/*
'
'
'
' Example:
'
' CreateListBoxHeader(Me.listBox_Body, Me.listBox_Header, Array("Header 1", "Header 2"))
'
'*/
Public Function CreateListBoxHeader(body As MSForms.listBox, header As MSForms.listBox, arrHeaders)
   ' make column count match
   header.ColumnCount = body.ColumnCount
   header.ColumnWidths = body.ColumnWidths

   ' add header elements
   header.Clear
   header.AddItem
   Dim i As Integer
   For i = 0 To UBound(arrHeaders)
      header.List(0, i) = arrHeaders(i)
   Next i

   ' make it pretty
   body.ZOrder (1)
   header.ZOrder (0)
   header.SpecialEffect = fmSpecialEffectFlat
   header.BackColor = RGB(200, 200, 200)
   header.Height = 10

   ' align header to body (should be done last!)
   header.Width = body.Width
   header.Left = body.Left
   header.Top = body.Top - (header.Height - 1)
End Function



'/*
'
'Size the ListBox Columns base on content
'For use create a hidden Label named lblHidden with Autosize: true and wordwarp:false
'
'*/
Public Function SetColumnWidthsAndHeader(ByRef body_listBox As MSForms.listBox, ByRef lblHidden As Variant, arrHeaders As Variant, header_listbox As MSForms.listBox)

   Dim i As Long, j As Long
   Dim sColWidths As String
   Dim dMax As Double


   If (body_listBox.ColumnCount) Then
      For i = 0 To body_listBox.ColumnCount - 1
         For j = 0 To body_listBox.ListCount - 1
            lblHidden.Caption = body_listBox.column(i, j) & "MM"
            If dMax < lblHidden.Width Then
               dMax = lblHidden.Width
            End If
         Next j

If (i <= UBound(arrHeaders)) Then
         lblHidden.Caption = arrHeaders(i) & "MM"
         If (dMax < lblHidden.Width) Then
            dMax = lblHidden.Width
         End If
End If
         sColWidths = sColWidths & CLng(dMax) & ";"
         dMax = 0
      Next i

      body_listBox.ColumnWidths = sColWidths
      lblHidden.Caption = ""

   End If


   Call Xform.CreateListBoxHeader(body_listBox, header_listbox, arrHeaders)

   header_listbox.Visible = True
End Function
