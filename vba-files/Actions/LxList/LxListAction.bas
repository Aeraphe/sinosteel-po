Attribute VB_Name = "LxListAction"


'namespace=vba-files\Actions\LxList


Const ROMANEIO_MAP_FIRST_LINE = 8
Const LX_FISRT_ROW = 9


Public Function addSelectedItemFormLxToRomaneio()

  LxSheet.Activate
  Dim lxSelectedRow As Integer
  lxSelectedRow = Application.ActiveCell.row
  Dim lxSelectedColumn As Integer
  lxSelectedColumn = Application.ActiveCell.column

  Dim romId As Integer

  romId = ConfigSheet.Range("CONFIG_ROMANEIO_ID").Value



  Dim romaneioLine As Integer
  Dim cwp As String



  If (lxSelectedRow >= LX_FISRT_ROW) Then

    RomaneioMapSheet.Activate

    Dim LastRow As Long
    With ActiveSheet

      romaneioLine = .Cells(.Rows.count, "A").End(xlUp).row + 1
      If (romaneioLine = 7) Then
        romaneioLine = ROMANEIO_MAP_FIRST_LINE
      End If
    End With

    Call delayMs(1000)
    cwp = LxSheet.Range("F" & lxSelectedRow).Value 'CWP
    RomaneioMapSheet.Range("A" & romaneioLine).Select
    RomaneioMapSheet.Range("A" & romaneioLine).Value = romId + 1
    RomaneioMapSheet.Range("F" & romaneioLine).Value = LxSheet.Range("H" & lxSelectedRow).Value  'Pos
    RomaneioMapSheet.Range("L" & romaneioLine).Value = LxSheet.Range("I" & lxSelectedRow).Value  'Descri
    RomaneioMapSheet.Range("M" & romaneioLine).Value = LxSheet.Range("N" & lxSelectedRow).Value  'Des.
    RomaneioMapSheet.Range("N" & romaneioLine).Value = LxSheet.Range("O" & lxSelectedRow).Value  'Des.rev
    RomaneioMapSheet.Range("O" & romaneioLine).Value = LxSheet.Range("G" & lxSelectedRow).Value  'Tag
    RomaneioMapSheet.Range("G" & romaneioLine).Value = cwp
    RomaneioMapSheet.Range("H" & romaneioLine).Value = LxSheet.Range("K" & lxSelectedRow).Value  'Unidade

    RomaneioMapSheet.Range("I" & romaneioLine).Value = LxSheet.Range("J" & lxSelectedRow).Value  'Qt
    RomaneioMapSheet.Range("J" & romaneioLine).Value = LxSheet.Range("L" & lxSelectedRow).Value  'Peso
    RomaneioMapSheet.Range("K" & romaneioLine).Value = LxSheet.Range("M" & lxSelectedRow).Value  'Peso total
    RomaneioMapSheet.Range("AH" & romaneioLine).Value = LxSheet.Range("A" & lxSelectedRow).Value  'ID Mat

    RomaneioMapSheet.Range("T" & romaneioLine).Value = SearchCWPColor(cwp)
    Call delayMs(500)
    LxSheet.Range("S" & lxSelectedRow).Value = "Adicionado"


    ConfigSheet.Range("CONFIG_ROMANEIO_ID").Value = romId + 1
  End If


  LxSheet.Activate
  'Move tothe next Filtered Cell
  ActiveCell.offset(1, 0).Select
  Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.offset(1, 0).Select
  Loop


End Function


Private Sub delayMs(ms As Long)
  Debug.Print TimeValue(Now)
  Application.Wait (Now + (ms * 0.00000001))
  Debug.Print TimeValue(Now)
End Sub



Private Function SearchCWPColor(cwp As String) As String

  Dim line  As Integer
  line = 6


  While line <= 11

    If (CWPSheet.Cells(line, "B").Value = cwp) Then

      SearchCWPColor = CWPSheet.Cells(line, "C").Value
      Exit Function
    End If

    line = line + 1
  Wend

End Function
