Attribute VB_Name = "SupplierController"


'namespace=vba-files\Controllers

Public Function Create()

    Dim itemRange As Range

    Dim response As Variant
    Set response = CreateObject("Scripting.Dictionary")
    Dim responseItem As Integer
    responseItem = 0


    For Each itemRange In AddSuppliersSheet.Range("ADD_SUPPLIERS_TABLE").Rows


        If (AddSuppliersSheet.Cells(itemRange.row, "A") <> "") Then

            Dim data As Variant
            Set data = CreateObject("Scripting.Dictionary")

            data("name") = AddSuppliersSheet.Cells(itemRange.row, "A").Value
            data("type") = AddSuppliersSheet.Cells(itemRange.row, "B").Value
            data("email") = AddSuppliersSheet.Cells(itemRange.row, "C").Value
            data("phone") = AddSuppliersSheet.Cells(itemRange.row, "D").Value
            data("address") = AddSuppliersSheet.Cells(itemRange.row, "E").Value

            Call SuppliersDataBase.Create(data)

        End If

    Next itemRange

End Function




