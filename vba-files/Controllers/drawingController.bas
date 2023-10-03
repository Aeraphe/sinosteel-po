Attribute VB_Name = "drawingController"


'namespace=vba-files\Controllers

Public Function AddDrawings()

    AddDrawingSheet.Activate
    
    Dim drawItemRange As Range

    Dim response As Variant
    Set response = CreateObject("Scripting.Dictionary")
    Dim responseItem As Integer
    responseItem = 0


    For Each drawItemRange In AddDrawingSheet.Range("ADD_DRAWING_TABLE").Rows


        If (AddDrawingSheet.Cells(drawItemRange.row, "A") <> "") Then
           
            Dim data As Variant
            Set data = CreateObject("Scripting.Dictionary")

            data("code") = AddDrawingSheet.Cells(drawItemRange.row, "A").Value
            data("rev") = AddDrawingSheet.Cells(drawItemRange.row, "B").Value
            data("tag") = AddDrawingSheet.Cells(drawItemRange.row, "C").Value
            data("name") = AddDrawingSheet.Cells(drawItemRange.row, "D").Value
            data("description") = AddDrawingSheet.Cells(drawItemRange.row, "E").Value
            data("weight") = AddDrawingSheet.Cells(drawItemRange.row, "F").Value


          
           Call DrawingsDataBase.addDrawing(data)
        
        End If

    Next drawItemRange

    AddDrawingSheet.Range("ADD_DRAWING_TABLE").ClearContents
End Function
