Attribute VB_Name = "act_comments_notifi"


'namespace=vba-files\Actions




'/*
' This function creates and sends an email to recipients.
' The input parameters include:
' - request_id: the ID of the request
' - Outlook_App: the Outlook application object
' - attachs (optional): attachments for the email
'*\
Function make(ByVal projectName As String, ByVal documents As Object, Outlook_App As Object, Optional attachs As Object)


    ' Variables for recipient_id, project_id, emails, HTML, and full_title.
    Dim recipient_id As String
    Dim project_id As String
    Dim emails As Object
    Dim HTML As Object
    Dim full_title As String

    ' Create a dictionary object for emails.
    Set emails = CreateObject("Scripting.Dictionary")

    ' Create the HTML body of the email.
     Set HTML = makeHtmlBody(documents)

    ' Get the email addresses of the recipients.
    Set emails = get_emails(1)

    ' Create the full title for the email.
    full_title = UCase(projectName) & " - NOTIFICAÇÃO DE DOCUMENTOS COMENTADOS - " & Now()

    ' Call the service_email.Send method to send the email.
    Call service_email.Send(emails, full_title, HTML("HTML"), Outlook_App)
End Function




Private Function makeHtmlBody(ByVal documents As Object) As Object

    Const HTML_BR = "<br><br><br>"
    Const HTML_BODY = "<body>"
    Const HTML_CLOSE = "</body></html>"

    Dim response As Object
    Set response = CreateObject("Scripting.Dictionary")

    Dim emailMsgConf As ADODB.Recordset
    Set emailMsgConf = db_email.get_layout("NOTFI_COMMENTS")

    Dim htmlHeader As String
    htmlHeader = XdbFactory.getData(emailMsgConf, "html_header")

    Dim headerCssMsg As String
    headerCssMsg = XdbFactory.getData(emailMsgConf, "header_css_msg")


    Dim midleMsg As String
    midleMsg = XdbFactory.getData(emailMsgConf, "midle_msg")

    Dim posMsg As String
    posMsg = XdbFactory.getData(emailMsgConf, "pos_msg")

    Dim table As String
    table = Replace(midleMsg, "[R0]", "DOCUMENTOS COMENTADOS")
    table = Replace(table, "[R1]", Now())

    Dim htmlTr As String
    htmlTr = XdbFactory.getData(emailMsgConf, "html_tr")

    Dim htmlDocTable As Object

    Set htmlDocTable = makeHtmlDocTb(documents, table, htmlTr)

    Dim HTML_WELCOME_MSG As String
    HTML_WELCOME_MSG = make_welcome_msg()


    Dim preMsg As String
    preMsg = XdbFactory.getData(emailMsgConf, "pre_msg")

    Dim countFiles As String
    countFiles = helper_string.StringFormat(preMsg, "Total de Documentos: " & htmlDocTable("COUNT"))


    Dim htmlResponse As String
    htmlResponse = htmlHeader & HTML_BODY & HTML_WELCOME_MSG & countFiles & htmlDocTable("TABLE") & posMsg & HTML_BR & HTML_BR & HTML_CLOSE

    response("HTML") = htmlResponse

    Set makeHtmlBody = response
End Function

'/*
'
'Make a HTML Welcome Message
'
'*/
Private Function make_welcome_msg() As String

    Dim response As String

    If TimeValue(Now) < TimeValue("13:00:00") Then
        response = "Prezados, Bom Dia!"
    ElseIf (TimeValue(Now) > TimeValue("13:00:00") And TimeValue(Now) < TimeValue("18:00:00")) Then
        response = "Prezados, Bom Tarde!"
    Else
        response = "Prezados, Boa Noite!"
    End If

    make_welcome_msg = helper_string.StringFormat("<p>{0}</p>", response)

End Function


'/*
'
'Make a HTML Table with Documents
'
'*/
Private Function makeHtmlDocTb(ByRef documents As Object, ByVal HTML_TABLE As String, ByVal HTML_TR As String) As Object



    Dim HTML_TR_CONTENT As String
    Dim HTML_BODY As String
    Dim doc As String
    Dim sb As Object


    Set sb = New StringBuilder


    Dim count As Long
    count = 0
    For Each varKey In documents.Keys()



        doc = "<b>" & documents(varKey)("DOC") & "_Rev_" & documents(varKey)("REVIEW") & "_" & documents(varKey)("ISSUE") & "</b>  -->  [PROXIMA REV:" & documents(varKey)("NEXT_REVIEW") & "]   [PROXIMA TE:" & documents(varKey)("NEXT_ISSUE") & "]    -  " & documents(varKey)("DOC_INFO")
        HTML_TR_CONTENT = Replace(HTML_TR, "[R0]", doc)
        Call sb.AppendLine(HTML_TR_CONTENT)
    
        count = count + 1
    Next '

    Dim full_string  As String
    full_string = sb.ToString()

    Dim response As Object
    Set response = CreateObject("Scripting.Dictionary")

    response("TABLE") = Replace(HTML_TABLE, "[R2]", full_string)
    response("COUNT") = count

    Set makeHtmlDocTb = response

End Function


