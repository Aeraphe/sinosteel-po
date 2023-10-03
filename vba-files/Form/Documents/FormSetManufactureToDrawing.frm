VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSetManufactureToDrawing 
   Caption         =   "Definir Fabricante ao Desenho"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15405
   OleObjectBlob   =   "FormSetManufactureToDrawing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSetManufactureToDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents






Private Sub CommandButton1_Click()

   Dim respQuery As ADODB.Recordset
   Set respQuery = DrawingsDataBase.getAll(SearchText.Value)

   DrawingsList.Clear

   Do Until respQuery.EOF

      DrawingsList.AddItem XdbFactory.getData(respQuery, "id")

      DrawingsList.List(DrawingsList.ListCount - 1, 1) = XdbFactory.getData(respQuery, "supplier_number")
      DrawingsList.List(DrawingsList.ListCount - 1, 2) = XdbFactory.getData(respQuery, "equipament_tag")
      DrawingsList.List(DrawingsList.ListCount - 1, 3) = XdbFactory.getData(respQuery, "name")
      DrawingsList.List(DrawingsList.ListCount - 1, 4) = XdbFactory.getData(respQuery, "description")
      DrawingsList.List(DrawingsList.ListCount - 1, 5) = XdbFactory.getData(respQuery, "weight")

      respQuery.MoveNext
   Loop

End Sub






Private Sub DrawingsList_Change()

   getManufacturesHandler
End Sub


Private Function getManufacturesHandler()


   Dim respQuery As ADODB.Recordset
   Set respQuery = SuppliersDataBase.getAllSuppliersFromDrawing(DrawingsList.Value)


   ManufactureListBox.Clear

   Do Until respQuery.EOF

      ManufactureListBox.AddItem XdbFactory.getData(respQuery, "id")
      ManufactureListBox.List(ManufactureListBox.ListCount - 1, 1) = XdbFactory.getData(respQuery, "name")


      respQuery.MoveNext
   Loop
End Function


Private Sub SearchDescriptionbtn_Click()
   Dim respQuery As ADODB.Recordset
   Set respQuery = DrawingsDataBase.getAll(SearchText.Value, "description")

   DrawingsList.Clear

   Do Until respQuery.EOF

      DrawingsList.AddItem XdbFactory.getData(respQuery, "id")

      DrawingsList.List(DrawingsList.ListCount - 1, 1) = XdbFactory.getData(respQuery, "supplier_number")
      DrawingsList.List(DrawingsList.ListCount - 1, 2) = XdbFactory.getData(respQuery, "equipament_tag")
      DrawingsList.List(DrawingsList.ListCount - 1, 3) = XdbFactory.getData(respQuery, "name")
      DrawingsList.List(DrawingsList.ListCount - 1, 4) = XdbFactory.getData(respQuery, "description")
      DrawingsList.List(DrawingsList.ListCount - 1, 5) = XdbFactory.getData(respQuery, "weight")

      respQuery.MoveNext
   Loop
End Sub

Private Sub UserForm_Activate()


   Dim respQuery As ADODB.Recordset
   Set respQuery = SuppliersDataBase.getAll()

   ManufacureCb.Clear

   Do Until respQuery.EOF

      ManufacureCb.AddItem XdbFactory.getData(respQuery, "id") & " - " & XdbFactory.getData(respQuery, "name")


      respQuery.MoveNext
   Loop
End Sub




Private Sub AddManufacture_Click()

   Dim data As Variant
   Set data = CreateObject("Scripting.Dictionary")

   Dim arrSplitStrings1 As Variant
   arrSplitStrings1 = Split(ManufacureCb.Value, " - ")

   If (DrawingsList.Value <> "" And arrSplitStrings1(0) <> "") Then
      data("drawing_id") = DrawingsList.Value
      data("manufactor_id") = arrSplitStrings1(0)



      Call SuppliersDataBase.SetManufactureToDrawing(data)

      getManufacturesHandler

   End If
End Sub


Private Sub DeleteManufacture_Click()
   Dim data As Variant
   Set data = CreateObject("Scripting.Dictionary")
   If (DrawingsList.Value <> "") Then

      data("drawing_id") = DrawingsList.Value
      data("manufactor_id") = ManufactureListBox.Value
      Call SuppliersDataBase.RemoveMAnufacturerFromDrawing(data)
      getManufacturesHandler
   End If
End Sub
