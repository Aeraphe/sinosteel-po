Attribute VB_Name = "act_doc_request"


'namespace=vba-files\Actions\DocRequestAction



Public Function changeDocRequestStatus(ByVal docRequestId As String, ByVal status As String)



    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
 
    data("status") = status
    data("status_date") = format(Now(), "yyyy-mm-dd hh:mm:ss")
    data("post_user_response_id") = auth.get_user_id
    Call db_issue_request.updateRequestDocument(data, docRequestId)
 
 End Function
 

    
Public Function changeDocRequestStatus2(ByVal docRequestId As String, ByVal status As String)



    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
 
    data("status") = status
    data("status_date") = format(Now(), "yyyy-mm-dd hh:mm:ss")
    Call db_issue_request.updateRequestDocument(data, docRequestId)
 
 End Function
