VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grd_simple_confirmatiom_form 
   Caption         =   "Qual GRD Quer Gerar?"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   OleObjectBlob   =   "grd_simple_confirmatiom_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grd_simple_confirmatiom_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\GRD






Public options As Object
Public confirmation As Boolean

Private Sub save_btn_Click()
 
   answer = MsgBox("Tem certeza que quer criar a GRD?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer) Then

   Set options = CreateObject("Scripting.Dictionary")
   confirmation = True
   
   
  options("GRD_FILE") = grd_ck.Value
  options("GRD_SUPPLIER") = grd_supplier_ck.Value
  options("EMAIL") = email_ck.Value
  
  Me.Hide
 Else
  confirmation = False
 End If
 
 

End Sub
