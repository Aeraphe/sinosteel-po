VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} grd_create_confirmation_form 
   Caption         =   "Deseja gerar a GRD?"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16005
   OleObjectBlob   =   "grd_create_confirmation_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "grd_create_confirmation_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'namespace=vba-files\Form\GRD



Public confirmation As Boolean
Public projectIdSelected As String
Public docList As Object

Private Sub UserForm_Initialize()
 confirmation = False

 populate_destiny_select_handler
 
 select_satus.AddItem "For Reviewd"
 select_satus.AddItem "For Reviewd"
 
 select_satus.ListIndex = 0
 
 grd_description_txt.Value = auth.user_name & "_" & Now

End Sub

Private Function populate_destiny_select_handler()

   Dim respQuery As ADODB.Recordset
   Set respQuery = db_grd.getAll()
   Call Shared_CommonSelectComp.Mount(grd_to_select, respQuery, "id", "name")

End Function

Private Sub close_btn_Click()
 confirmation = False
  Me.Hide
End Sub

Private Sub save_btn_Click()


   Dim answer As Integer
   Dim grd_id  As Long
   Dim options As Object
   
   Set options = CreateObject("Scripting.Dictionary")

   answer = MsgBox("Tem certeza que quer criar a GRD?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")
   If (answer) Then
   
   Load grd_create_form

   grd_create_form.project_txt.Value = project_txt.Value
   grd_create_form.projectIdSelected = projectIdSelected
   grd_create_form.grd_to_select.Value = grd_to_select.Value
   grd_create_form.grd_description_txt.Value = grd_description_txt.Value
   grd_create_form.grd_obs.Value = grd_obs_txt.Value
   

   Call loadDocsInGRD(grd_create_form.grd_list)



  options("GRD_FILE") = grd_ck.Value
  options("GRD_SUPPLIER") = grd_supplier_ck.Value
  options("EMAIL") = email_ck.Value
  
 Load UserFormAlert
 UserFormAlert.Label1.Caption = "GERANDO A GRD"
 UserFormAlert.Show
  Call grd_create_form.GenerateGRDHandler(options)
  
  Unload grd_create_form
 
  Unload Me
End If

End Sub


Private Function loadDocsInGRD(ByRef grdList As Object)

Dim doc As Object

    For Each varKey In docList.Keys()
    If (varKey <> "") Then
       Set doc = docList(varKey)
       
       
            grdList.AddItem doc("docRevId")
            grdList.List(grdList.ListCount - 1, 1) = UCase(doc("docNumber"))
            grdList.List(grdList.ListCount - 1, 2) = "[Rev: " & doc("docNextRev") & "]   [TE: " & issue & "]"
            grdList.List(grdList.ListCount - 1, 3) = Left(doc("desciption"), 160)
            grdList.List(grdList.ListCount - 1, 4) = doc("category")
            grdList.List(grdList.ListCount - 1, 5) = doc("media")
            grdList.List(grdList.ListCount - 1, 6) = doc("type")
            grdList.List(grdList.ListCount - 1, 7) = doc("copies")
            grdList.List(grdList.ListCount - 1, 8) = doc("pages")
    End If
    Next
    
End Function
