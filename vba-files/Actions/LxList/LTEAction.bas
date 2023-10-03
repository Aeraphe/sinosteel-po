Attribute VB_Name = "LTEAction"


'namespace=vba-files\Actions\LxList

Const MAPA_ROMANEIOS_FIRST_LINE = 8

Public Function composeLTE()


    LTESheet.Activate

    LTESheet.Range("LTE_ITEMS_TABLE").Clear
    Dim LTENumber As String

    LTENumber = LTESheet.Range("LTE_N").Value

    Dim mapaLine As Long
    mapaLine = MAPA_ROMANEIOS_FIRST_LINE

    Dim item As Integer
    item = 1

    Dim pesoUnit As Variant
    pesoUnit = 0
    Dim quantity As Variant
    quantity = 0
    While RomaneioMapSheet.Cells(mapaLine, "A").Value <> ""

        If (RomaneioMapSheet.Cells(mapaLine, "C").Value = LTENumber) Then

            LTESheet.Range("FOR_NOME").Value = RomaneioMapSheet.Cells(mapaLine, "D").Value 'Fornecedor/Empresa
            LTESheet.Range("TRANSP").Value = RomaneioMapSheet.Cells(mapaLine, "AA").Value 'Fornecedor/Empresa
            LTESheet.Range("FOR_CWP").Value = RomaneioMapSheet.Cells(mapaLine, "G").Value 'CWP
            LTESheet.Range("DATA_EMB").Value = RomaneioMapSheet.Cells(mapaLine, "AB").Value 'Data Embarque
            LTESheet.Range("DATA").Value = RomaneioMapSheet.Cells(mapaLine, "AC").Value 'Data Recebido
            LTESheet.Range("NF").Value = RomaneioMapSheet.Cells(mapaLine, "Y").Value 'N nota
            LTESheet.Range("DATA_EM").Value = RomaneioMapSheet.Cells(mapaLine, "Z").Value 'Data nota
            LTESheet.Range("RECEBIDO_POR").Value = RomaneioMapSheet.Cells(mapaLine, "E").Value 'Data nota
            LTESheet.Cells(item + 21, "A").Value = item

            LTESheet.Cells(item + 21, "B").Value = RomaneioMapSheet.Cells(mapaLine, "H").Value 'Unidade
            LTESheet.Cells(item + 21, "C").Value = RomaneioMapSheet.Cells(mapaLine, "L").Value 'Descricao
            LTESheet.Cells(item + 21, "D").Value = RomaneioMapSheet.Cells(mapaLine, "F").Value 'Cod.MAt
            LTESheet.Cells(item + 21, "E").Value = RomaneioMapSheet.Cells(mapaLine, "M").Value 'Des.
            LTESheet.Cells(item + 21, "F").Value = RomaneioMapSheet.Cells(mapaLine, "N").Value 'Des. Rev
            LTESheet.Cells(item + 21, "G").Value = RomaneioMapSheet.Cells(mapaLine, "O").Value 'Pos
            quantity = RomaneioMapSheet.Cells(mapaLine, "I").Value 'Qt
            LTESheet.Cells(item + 21, "H").Value = quantity

            pesoUnit = RomaneioMapSheet.Cells(mapaLine, "J").Value 'Peso unitario
            LTESheet.Cells(item + 21, "I").Value = pesoUnit
            If (VarType(pesoUnit) = vbLong And VarType(quantity) = vbLong) Then
                LTESheet.Cells(item + 21, "J").Value = pesoUnit * quantity
            Else
                LTESheet.Cells(item + 21, "J").Value = "-"
            End If

            LTESheet.Cells(item + 21, "K").Value = RomaneioMapSheet.Cells(mapaLine, "P").Value 'Origin
            LTESheet.Cells(item + 21, "L").Value = RomaneioMapSheet.Cells(mapaLine, "Q").Value 'Area de Armazenagem
            LTESheet.Cells(item + 21, "M").Value = RomaneioMapSheet.Cells(mapaLine, "R").Value 'Tipo Embalagem
            LTESheet.Cells(item + 21, "N").Value = RomaneioMapSheet.Cells(mapaLine, "S").Value & " - " & RomaneioMapSheet.Cells(mapaLine, "R").Value & " - " & RomaneioMapSheet.Cells(mapaLine, "T").Value 'Vol
            LTESheet.Cells(item + 21, "O").Value = RomaneioMapSheet.Cells(mapaLine, "U").Value & " x " & RomaneioMapSheet.Cells(mapaLine, "V").Value & " x " & RomaneioMapSheet.Cells(mapaLine, "W").Value 'Vol

            item = item + 1
        End If
        mapaLine = mapaLine + 1
    Wend

    'Format Table
    LTESheet.Range("LTE_ITEMS_TABLE").Select

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("22:22").EntireRow.AutoFit
End Function






Public Function MasCreateLTEFiles(printFile As Boolean)

    Dim line As Integer
    line = 5

    While MassLTECreateSheet.Cells(line, "A") <> ""


        LTESheet.Range("LTE_N").Value = MassLTECreateSheet.Cells(line, "A")
        composeLTE
        createLTEFile
        Call delayMs(1000)
        If (printFile) Then
            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
            Call delayMs(7000)
        End If
        line = line + 1
    Wend

End Function


Private Sub delayMs(ms As Long)
    Debug.Print TimeValue(Now)
    Application.Wait (Now + (ms * 0.00000001))
    Debug.Print TimeValue(Now)
End Sub

