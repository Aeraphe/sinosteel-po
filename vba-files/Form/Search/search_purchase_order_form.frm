VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} search_purchase_order_form 
   Caption         =   "Search Purchase Order"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10050
   OleObjectBlob   =   "search_purchase_order_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "search_purchase_order_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Search




Private Sub search_project_btn_Click()
   Dim data As Object
   Set data = CreateObject("Scripting.Dictionary")
   Set data = Shared_ProjectSelectComp.GetProjetSelectedInForm

   Call Shared_PurchaseOrderSelectComp.Mount(purchase_orders_list, data("id"))

End Sub
