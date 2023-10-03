Attribute VB_Name = "helper_review"




'namespace=vba-files\Helpers




Function check_review(doc_id As String, last_rev As Variant, next_rev As Variant) As Object



 If (is_review_already_exist(doc_id, next_rev)) Then
 Set check_review = is_next_review_valid(last_rev, next_rev)
 
 Exit Function
 End If
 

    Dim response As Object
    Set response = CreateObject("Scripting.Dictionary")
    
       response("status") = False
       response("type") = "REV_ERROR"
       response("msg") = "A Revisão do documento já Existe"
       
       
       Set check_review = response
       
End Function


Function is_next_review_valid(last_rev As Variant, next_rev As Variant) As Object


    Dim response As Object
    Set response = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    last_rev = CInt(last_rev)

    On Error Resume Next
    next_rev = CInt(next_rev)

If (last_rev = -1) Then

If (next_rev = 0 Or next_rev = "A") Then
            response("status") = True
            response("type") = "REV_VALIDA"
            response("msg") = "Revisão está correta"
            Set is_next_review_valid = response
            Exit Function
            
            Else
                response("status") = False
                response("type") = "REV_ERROR"
                response("msg") = "A proxima revisão tem que ser igual a ZERO Ou A"
                Set is_next_review_valid = response
                Exit Function
            
End If
Else
    If (IsNumeric(last_rev)) Then

        If (IsNumeric(next_rev)) Then

            If (next_rev = last_rev + 1) Then
            
                response("status") = True
                response("type") = "REV_VALIDA"
                response("msg") = "A próxima Revisão ( " & next_rev & " ) está correta"
                Set is_next_review_valid = response
                Exit Function

            Else

                response("status") = False
                response("type") = "REV_ERROR"
                response("msg") = "A proxima revisão tem que ser igual a " & last_rev + 1
                Set is_next_review_valid = response
                Exit Function

            End If

Else
                response("status") = False
                response("type") = "REV_ERROR"
                response("msg") = "A proxima revisão tem que ser igual a  " & last_rev + 1
                Set is_next_review_valid = response
                Exit Function
        End If

    Else


        If (Not IsNumeric(next_rev)) Then

            last_rev_asc = Asc(last_rev)
            next_rev_asc = Asc(next_rev)

            If (last_rev_asc < next_rev_asc And next_rev_asc = last_rev_asc + 1) Then

                response("status") = True
                response("type") = "REV_VALIDA"
                response("msg") = "Revisão está correta"
                Set is_next_review_valid = response
                Exit Function
            Else
                response("status") = False
                response("type") = "REV_ERROR"
                response("msg") = "A proxima revisão tem que ser igual a " & CStr(last_rev_asc + 1) & " ou igual a 0 (ZERO)"
                Set is_next_review_valid = response
                Exit Function

            End If


        ElseIf (IsNumeric(next_rev) And next_rev = 0) Then
            response("status") = True
            response("type") = "REV_VALIDA"
            response("msg") = "Revisão está correta"
            Set is_next_review_valid = response
            Exit Function

        Else

            response("status") = False
            response("type") = "REV_ERROR"
            response("msg") = "A Revisão deve ser um número e igual a 0 (ZERO)"
            Set is_next_review_valid = response
            Exit Function

        End If
        End If
    End If


End Function


Private Function is_review_already_exist(doc_id As String, next_rev As Variant) As Boolean


      Dim respQuery As ADODB.Recordset
      Set respQuery = db_documents.SearchReviews(doc_id)

      Do Until respQuery.EOF

         rev = XdbFactory.getData(respQuery, "rev_code")

         If (next_rev = rev) Then

            is_review_already_exist = False
            
            Exit Function
         End If

         respQuery.MoveNext

      Loop
      
      is_review_already_exist = True
End Function
