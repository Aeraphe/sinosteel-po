Attribute VB_Name = "act_doc_change_status"


'namespace=vba-files\Actions\Documents






Public Function ConfirmDocumentStatus(ByVal status As String) As Boolean
    Dim answer As VbMsgBoxResult
    Dim msg As String

    Select Case status
     Case Constants.CONCLUIDO
        msg = "Tem serteza que foi confirmado o recebimento Do(s) documento(s) ?" & vbNewLine & " Stauts: CONCLUIDO"
     Case Constants.REJEITADO
        msg = "Quer rejeitar o(s) Documento(s)?" & vbNewLine & " Stauts: REJEITADO"
     Case Constants.ENVIADO
        msg = "Tem certeza que o documento já foi enviado para aprovação?" & vbNewLine & " Stauts: ENVIADO"
     Case Else
        msg = "Tem certeza que o deseja mudar o status Do(s) documento(s) para: " & status & "?"
    End Select

    answer = MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

    If answer = vbYes Then
        ConfirmDocumentStatus = True
    Else
        ConfirmDocumentStatus = False
    End If
End Function








'/*
'
' This Function changes the status of a document request With the given ID To a New status value.
' It returns a string message indicating the result of the status change operation.
'
'payload("INFO")
'payload("SCHEDULE_DATE")
'payload("SCHEDULE_INFO")
'payload("USER_OWNER_ID")
'
'@return resposne: Object Dictionary
'
'  resposne("INFO") As string
'  resposne("STATUS") As Boolean
'
'*/
Public Function ChangeStatus(ByVal currentStatus As String, ByVal newStatus As String, ByVal docRequestId As String, ByVal payload As Object) As Object

    Dim resposne As Object
    Set resposne = CreateObject("Scripting.Dictionary")
    Dim isDocumentStatusChange As Boolean
    ' Check If the user has permission To change the document request status.
    If Not CanUserChangeDocRequestStatus(docRequestId) Then
        resposne("INFO") = "Usuário não tem permissão."
        resposne("STATUS") = False

        Set ChangeStatus = resposne
     Exit Function
    End If

    ' Check If the New status value is valid And can be updated from the current status.
    Dim validUpdate As Boolean
    validUpdate = CanUpdateStatus(currentStatus, newStatus)

    If validUpdate Then
        ' Consider processing document status update asynchronously in a separate thread.
        ' Update the document status With the New status value And send a message (If provided).
        isDocumentStatusChange = ProcessDocumentStatus(currentStatus, newStatus, docRequestId, payload)

        If (isDocumentStatusChange) Then
            ' Set the return value To indicate success And include the New status value in the message.

            resposne("INFO") = "Status Modificado para: {" & newStatus & "}"
            resposne("STATUS") = True
            Set ChangeStatus = resposne
         Exit Function
        Else

            resposne("INFO") = "Mudança de Status Não Permitida: {" & newStatus & "}"
            resposne("STATUS") = True
            Set ChangeStatus = resposne
        End If
    Else
        ' Set the return value To indicate failure.
        resposne("INFO") = "Mudança de Status Não Permitida"
        resposne("STATUS") = False
        Set ChangeStatus = resposne
     Exit Function
    End If

End Function



Private Function CanUserChangeDocRequestStatus(ByVal docRequestId As String) As Boolean
    Dim useRequestOwnerId As Variant

    ' Get user id associated With the document request
    useRequestOwnerId = XdbFactory.getData(db_issue_request.getDocRequestData(docRequestId), "user_id")

    ' Check If current user is authorized To change the document request status, And return result
    CanUserChangeDocRequestStatus = auth.is_authorized("SUPER_ADMIN") Or auth.get_user_id() = useRequestOwnerId
End Function



Function CanUpdateStatus(ByVal currentStatus As String, ByVal newStatus As String) As Boolean

    On Error GoTo ErrorHandler

        ' Define allowed transitions between status
        Static allowedTransitions As Object
            If allowedTransitions Is Nothing Then
                Set allowedTransitions = CreateObject("Scripting.Dictionary")
                allowedTransitions.Add Constants.EMITIR, Array(Constants.EMITIR, Constants.PROGRAMADO, Constants.NO_FLUXO, Constants.CANCELADO, Constants.REJEITADO, Constants.HOLD, Constants.PEND)
                allowedTransitions.Add Constants.CANCELADO, Array(Constants.CANCELADO, Constants.EMITIR, Constants.PROGRAMADO, Constants.NO_FLUXO, Constants.HOLD, Constants.PEND)
                allowedTransitions.Add Constants.PROGRAMADO, Array(Constants.PROGRAMADO, Constants.EMITIR, Constants.NO_FLUXO, Constants.CANCELADO, Constants.REJEITADO, Constants.HOLD, Constants.PEND)
                allowedTransitions.Add Constants.NO_FLUXO, Array(Constants.NO_FLUXO, Constants.ENVIADO, Constants.CANCELADO, Constants.REJEITADO, Constants.PROGRAMADO, Constants.EMITIR, Constants.HOLD, Constants.PEND)
                allowedTransitions.Add Constants.LIB_ENG, Array(Constants.LIB_ENG, Constants.NO_FLUXO, Constants.ENVIADO, Constants.CANCELADO, Constants.REJEITADO, Constants.PROGRAMADO, Constants.EMITIR, Constants.HOLD, Constants.PEND)
                allowedTransitions.Add Constants.ENVIADO, Array(Constants.PROGRAMADO, Constants.NO_FLUXO, Constants.CONCLUIDO, Constants.REJEITADO)
                allowedTransitions.Add Constants.CONCLUIDO, Array(Constants.PROGRAMADO, Constants.NO_FLUXO, Constants.ENVIADO, Constants.REJEITADO)
                allowedTransitions.Add Constants.PEND, Array(Constants.PEND, Constants.NO_FLUXO, Constants.EMITIR, Constants.PROGRAMADO, Constants.REJEITADO, Constants.CANCELADO, Constants.HOLD)
                allowedTransitions.Add Constants.HOLD, Array(Constants.HOLD, Constants.NO_FLUXO, Constants.EMITIR, Constants.PROGRAMADO, Constants.REJEITADO, Constants.CANCELADO, Constants.HOLD, Constants.PEND)
                allowedTransitions.Add Constants.SUBISTITUIR, Array(Constants.SUBISTITUIR, Constants.REJEITADO)
            End If

            ' Check If transition is allowed
            Dim possibleTransitions() As Variant
            possibleTransitions = allowedTransitions(currentStatus)

            If Not IsArray(possibleTransitions) Then
                MsgBox "Unexpected current status: " & currentStatus
             Exit Function
            End If

            Dim isAllowed As Boolean
            isAllowed = False
            For Each transition In possibleTransitions
                If StrComp(transition, newStatus, vbTextCompare) = 0 Then
                    isAllowed = True
                 Exit For
                End If
                Next

                CanUpdateStatus = isAllowed
             Exit Function

ErrorHandler:
                CanUpdateStatus = False

End Function


'/*
'
' This Function updates the status of a document request With the given ID To the New status value.
' The Function takes three parameters: `newStatus`, `docRequestId` And `info`.
' It handles any errors that occur during the update process And displays an error message.
' payload("INFO")
' payload("SCHEDULE_DATE")
' payload("SCHEDULE_INFO")
' payload("USER_OWNER_ID")
'*/
Private Function ProcessDocumentStatus(ByVal currentStatus As String, ByVal newStatus As String, ByVal docRequestId As String, ByVal payload As Object) As Boolean

    ' Define a data object To pass along With the status update.
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data.Add "status", newStatus
    data.Add "status_date", format(Now(), "yyyy-mm-dd hh:mm:ss")
    data.Add "obs", payload("INFO")

    Dim response As Boolean
    response = False

    ' Use a Select Case block To perform different actions based on the New status value.
    Select Case newStatus

     Case Constants.EMITIR:

        data.Add "post_in_date", ""
        data.Add "post_user_response_id", payload("USER_OWNER_ID")
        data.Add "post_user_response_msg", payload("INFO")

        response = SendAction(docRequestId, data)

     Case Constants.LIB_ENG, Constants.PEND, Constants.CANCELADO, Constants.HOLD:

        response = DefaultAction(docRequestId, data)

     Case Constants.NO_FLUXO:

        If (currentStatus = Constants.ENVIADO Or currentStatus = Constants.REJEITADO) Then
            response = DispatchAction(Constants.REVIEW_SATUS_SEND, docRequestId, data)
        Else
            response = DefaultAction(docRequestId, data)
        End If


     Case Constants.ENVIADO:

        response = DispatchAction(Constants.REVIEW_SATUS_EXP, docRequestId, data)

     Case Constants.CONCLUIDO:

        response = DispatchAction(Constants.REVIEW_SATUS_POST, docRequestId, data)

     Case Constants.REJEITADO

        response = RejectAction(docRequestId, data)

     Case Constants.PROGRAMADO:

        data.Add "post_in_date", payload("SCHEDULE_DATE")
        data.Add "post_user_response_id", payload("USER_OWNER_ID")
        data.Add "post_user_response_msg", payload("INFO")


        If (payload("GRD_ID") <> "") Then
            response = DispatchAction("LIB", docRequestId, data)
            If (MsgBox("Tem certeza que quer excluir da GRD o Documento: " & payload("DOC_NUMBER") & " Reprogramado?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")) Then
              Call db_grd.softDeleteDocument(payload("GRD_ID"), payload("DOC_REVIEW_ID"))
            End If
      
        Else
            response = DefaultAction(docRequestId, data)
        End If

    End Select


    ProcessDocumentStatus = response
 Exit Function

    ' Handle any errors that occur during the update process And display an error message.
ErrorHandler:
    ProcessDocumentStatus = False
    MsgBox "Error " & Err.Number & ": " & Err.description
End Function


Function DefaultAction(ByVal docRequestId As String, ByVal data As Object) As Boolean
    ' Error handling in Case an error occurs during the procedure
    On Error GoTo ErrorHandler
        Call db_issue_request.updateRequestDocument(data, docRequestId)
        DefaultAction = True
     Exit Function
ErrorHandler:
        DefaultAction = False
        ' Display error message If an error occurs during the procedure
        MsgBox "Error: " & Err.description

End Function


Function SendAction(ByVal docRequestId As String, ByVal data As Object) As Boolean
    ' Error handling in Case an error occurs during the procedure
    On Error GoTo ErrorHandler


        ' Declare variables
        Dim response As Boolean
        Dim respRequestQuery As ADODB.Recordset
        Dim grd As String

        response = False

        ' Retrieve document request data from the database using docRequestId As parameter
        Set respRequestQuery = db_issue_request.getDocRequestData(docRequestId)

        ' Get values from the recordset using XdbFactory.getData method
        grd = XdbFactory.getData(respRequestQuery, "grd_id")
        If (grd = "") Then
            Call db_issue_request.updateRequestDocument(data, docRequestId)
            response = True
        End If

        SendAction = response
     Exit Function
ErrorHandler:
        SendAction = response
        ' Display error message If an error occurs during the procedure
        MsgBox "Error: " & Err.description

End Function


Function DispatchAction(ByVal docReviewStatus As String, ByVal docRequestId As String, ByVal data As Object) As Boolean

    ' Error handling in Case an error occurs during the procedure
    On Error GoTo ErrorHandler

        ' Declare variables
        Dim respRequestQuery As ADODB.Recordset
        Dim docRevId As String

        ' Retrieve document request data from the database using docRequestId As parameter
        Set respRequestQuery = db_issue_request.getDocRequestData(docRequestId)

        ' Get values from the recordset using XdbFactory.getData method
        docRevId = XdbFactory.getData(respRequestQuery, "rev_id")

        If (docRevId <> "") Then

            Call db_issue_request.updateRequestDocument(data, docRequestId)
            Call db_documents.update_review_status(docRevId, docReviewStatus, Date)

            DispatchAction = True
        Else
            DispatchAction = False
        End If


     Exit Function
ErrorHandler:
        DispatchAction = False
        ' Display error message If an error occurs during the procedure
        MsgBox "Error: " & Err.description
End Function








'/*
' This Sub procedure rejects a document request For a specific motive.
' It retrieves data of the document request from the database, constructs a dictionary object containing
' details about the rejection, And calls a separate action To reject the document in ProjectWise.
'*/
Function RejectAction(ByVal docRequestId As String, ByVal data As Object) As Boolean

    ' Error handling in Case an error occurs during the procedure
    On Error GoTo ErrorHandler

        ' Declare variables
        Dim respRequestQuery As ADODB.Recordset
        Dim projectId As String
        Dim docRevId As String
        Dim docId As String
        Dim grdId As String
        Dim filePath As String
        Dim response As Boolean

        ' Retrieve document request data from the database using docRequestId As parameter
        Set respRequestQuery = db_issue_request.getDocRequestData(docRequestId)

        ' Get values from the recordset using XdbFactory.getData method
        projectId = XdbFactory.getData(respRequestQuery, "project_id")
        reviewCode = XdbFactory.getData(respRequestQuery, "rev_code_request")
        docRevId = XdbFactory.getData(respRequestQuery, "rev_id")
        docId = XdbFactory.getData(respRequestQuery, "id")
        grdId = XdbFactory.getData(respRequestQuery, "grd_id")
        filePath = XdbFactory.getData(respRequestQuery, "file_path")

        ' Close the recordset And Set it To Nothing To release memory
        respRequestQuery.Close
        Set respRequestQuery = Nothing

        ' Create a New Dictionary object To store rejection data
        Dim rejectData  As Object
        Set rejectData = CreateObject("Scripting.Dictionary")

        ' Add key-value pairs To the dictionary For various details related To the rejection
        rejectData.Add "motive", data("obs")
        rejectData.Add "doc_id", docId
        rejectData.Add "user_id", auth.get_user_id ' Get the user ID of the authenticated user
        rejectData.Add "review", reviewCode
        rejectData.Add "docReviewId", docRevId

        ' Call a separate action To reject the document in ProjectWise using projectId, filePath, And the rejection details
        response = action_reject_document.RejectDocument(projectId, filePath, rejectData)
        If (response) Then

            Call db_issue_request.updateRequestDocument(data, docRequestId)

        End If

        RejectAction = response
     Exit Function
ErrorHandler:
        ' Display error message If an error occurs during the procedure
        MsgBox "Error: " & Err.description
End Function



