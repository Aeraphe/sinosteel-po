Attribute VB_Name = "helper_grd_sandbox"


'namespace=vba-files\Helpers




Function Create(sandbox_name As String, grd_type As String) As Long

    If (Not IsNull(auth.get_user_id)) Then
        Dim sandbox_data As Object
        Set sandbox_data = CreateObject("Scripting.Dictionary")

        sandbox_data("user_id") = auth.get_user_id
        sandbox_data("name") = sandbox_name
        sandbox_data("grd_type") = grd_type

        Create = db_grd.create_sandbox(sandbox_data)
    End If
End Function


Public Function Insert(rev_id As Long, sandbox_id As Long)

    If (Not IsNull(rev_id) And Not IsNull(sandbox_id)) Then


        Dim data As Object
        Set data = CreateObject("Scripting.Dictionary")

        data("doc_review_id") = rev_id
        data("grd_sandbox_id") = sandbox_id

        Call db_grd.insert_doc_to_sandbox(data)
        
        End If

End Function



    
