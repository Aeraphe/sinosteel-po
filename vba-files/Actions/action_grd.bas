Attribute VB_Name = "action_grd"

'namespace=vba-files\Actions

Public Function create_selected_grd_view(opt As Object, grd_id As String)



   If (opt("GRD_FILE")) Then
      Call view_create_grd.publish(grd_id)
   End If

   If (opt("EMAIL")) Then
      Call helper_grd_email.make(grd_id)
   End If

   If (opt("GRD_SUPPLIER")) Then

      Call view_grd_vale.publish(grd_id)
   End If


End Function
