VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} doc_review_save_confirm_form 
   Caption         =   "Salvar no sistema as Revisões"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   OleObjectBlob   =   "doc_review_save_confirm_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "doc_review_save_confirm_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\Documents

Public confirmation As Boolean
Public options As Object


Private Sub cancel_btn_Click()
 
 confirmation = False
 Me.Hide
 
End Sub

Private Sub grd_create_now_opt_Click()

End Sub

Private Sub save_btn_Click()
  
  options("PROP") = Xhelper.iff(just_add_to_db_opt.Value, "JUST_ADD", Xhelper.iff(grd_create_now_opt.Value, "CREATE_GRD_NOW", ""))
 
  
  confirmation = True
 
 Me.Hide
End Sub

Private Sub UserForm_Activate()
 confirmation = False
 Set options = CreateObject("Scripting.Dictionary")

 
End Sub
