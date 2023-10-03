Attribute VB_Name = "helper_grd_email"

'namespace=vba-files\Helpers



Public Function make(ByVal grd_id As String)

    Dim recipient_id As String
    Dim project_id As String
    Dim emails As Object
    Dim Outlook_App As Object
    Dim HTML As Object
    Dim respGRD As ADODB.Recordset


    On Error Resume Next
    Set Outlook_App = GetObject(, "Outlook.Application")
    If Err.Number = 429 Then
        MsgBox "Erro: Gentileza abrir o Outlook Primeiro", vbCritical
    Else

        Set emails = CreateObject("Scripting.Dictionary")
        Set HTML = CreateObject("Scripting.Dictionary")
        Set respGRD = db_grd.getById(grd_id)

        recipient_id = XdbFactory.getData(respGRD, "recipent_id")
        project_id = XdbFactory.getData(respGRD, "project_id")
        grd_code = XdbFactory.getData(respGRD, "code")
        grd_date = XdbFactory.getData(respGRD, "issue_date")
        grd_sequence = XdbFactory.getData(respGRD, "sequece_number")
        email_msg_id = XdbFactory.getData(respGRD, "email_msg_id")
        grd_full_number = grd_code & grd_sequence


        Set HTML = make_html_body(project_id, grd_id, grd_full_number, grd_date, email_msg_id)
        Set emails = get_emails(recipient_id)
        title = HTML("PRE_TILE") & grd_full_number

        Call Send(emails, title, HTML("HTML"), Outlook_App)
    
    End If

End Function







Private Function make_html_body(ByVal project_id As String, ByVal grd_id As String, ByVal grd_full_number As String, ByVal grd_date As String, ByVal email_msg_id As String) As Object

    Dim response As Object
    Dim HTML_WELCOME_MSG As String
    Dim HTML  As String
    Dim HTML_HEADER  As String
    Dim HTML_BODY  As String
    Dim HTML_CLOSE  As String
    Dim HTML_RESPOSNSE As String
    Dim HTML_BR As String
    Dim HTML_DOC_TABLE As String
    Dim PRE_TITLE As String
    Dim emailMsgConf As ADODB.Recordset

    HTML_WELCOME_MSG = make_welcome_msg
    Set response = CreateObject("Scripting.Dictionary")

    HTML = config_sheet.Range("CONF_HTML").Value
    HTML_BR = "<br><br><br>"
    HTML_BODY = "<body>"
    HTML_CLOSE = "</body></html>"

    HTML_TOP_PART = HTML & HTML_HEADER & HTML_BODY & HTML_WELCOME_MSG


    Set emailMsgConf = db_email.get_config_msgs(project_id, email_msg_id)



    HTML_HEADER = XdbFactory.getData(emailMsgConf, "header_css_msg")
    PRE_TITLE = XdbFactory.getData(emailMsgConf, "pre_title")
    HTML_PRE_MSG = XdbFactory.getData(emailMsgConf, "pre_msg")
    HTML_MIDLE_MSG = XdbFactory.getData(emailMsgConf, "midle_msg")
    HTML_POS_MSG = XdbFactory.getData(emailMsgConf, "pos_msg")

    HTML_TABLE = Replace(HTML_MIDLE_MSG, "[R0]", grd_full_number)
    HTML_TABLE = Replace(HTML_TABLE, "[R1]", grd_date)

    HTML_DOC_TABLE = make_html_doc_tb(grd_id, HTML_TABLE)

    HTML_MIDLE_PART = HTML_PRE_MSG & HTML_DOC_TABLE & HTML_POS_MSG

    HTML_END_PART = HTML_BR & HTML_BR & HTML_CLOSE

    HTML_RESPOSNSE = HTML_TOP_PART & HTML_MIDLE_PART & HTML_END_PART

    response("HTML") = HTML_RESPOSNSE
    response("PRE_TILE") = PRE_TITLE

    Set make_html_body = response

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
    ElseIf (TimeValue(Now) > TimeValue("13:00:00") And TimeValue(Now) < TimeValue("17:00:00")) Then
        response = "Prezados, Bom Tarde!"
    Else

        response = "Prezados, Boa Noite!"
    End If

    make_welcome_msg = "<p>" & response & "</p>"
End Function

'/*
'
'Make a HTML Table with Documents
'
'*/
Private Function make_html_doc_tb(ByVal grd_id As String, ByVal HTML_TABLE As String) As String

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data("name") = grd_id

    Dim respGrdDocs As ADODB.Recordset
    Set respGrdDocs = db_grd.getGRDItems(data)

    Dim HTML_TD As String
    HTML_TD = config_sheet.Range("CONF_HTML_TD").Value


    Dim HTML_BODY As String
    Dim doc As String
    Dim doc_list As String

    Do Until respGrdDocs.EOF
        doc = XdbFactory.getData(respGrdDocs, "doc_number") & "_Rev_" & XdbFactory.getData(respGrdDocs, "rev_code") & " - " & XdbFactory.getData(respGrdDocs, "name") & " - " & XdbFactory.getData(respGrdDocs, "description")
        HTML_TD_CONTENT = Replace(HTML_TD, "[R0]", doc)
        doc_list = doc_list & HTML_TD_CONTENT & vbNewLine
        respGrdDocs.MoveNext
    Loop


    make_html_doc_tb = Replace(HTML_TABLE, "[R2]", doc_list)

End Function


'/*
'
'Get a HTML Signature from outlook
'
'*/
Private Function get_signature() As String
    Dim S As String
    S = Environ("appdata") & "\Microsoft\Signatures\"
    If Dir(S, vbDirectory) <> vbNullString Then S = S & Dir$(S & "*.htm") Else S = ""
        S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll


        get_signature = S

End Function


'/*
'
'Get all emails to send message
'
'*/
Private Function get_emails(ByVal recipient_id As String) As Object

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    Dim respEmails As ADODB.Recordset
    Set respEmails = db_grd_recipient.get_all_recipent_emails(recipient_id)


    Dim emails_to As String
    Dim emails_cc As String

    emails_cc = ""
    emails_to = ""

    Do Until respEmails.EOF
    email_type = XdbFactory.getData(respEmails, "type")
        If (email_type = "TO") Then
            emails_to = emails_to & XdbFactory.getData(respEmails, "email") & ";"
        End If
        If (email_type = "CC") Then
            emails_cc = emails_cc & XdbFactory.getData(respEmails, "email") & ";"
        End If
        respEmails.MoveNext
    Loop

    data("TO") = emails_to
    data("CC") = emails_cc

    Set get_emails = data
End Function




'/*
'
'Semd email message
'
'*/
Private Function Send(emails As Object, ByVal title As String, ByVal HTML_BODY As String, ByRef Outlook_App As Object, Optional action As String = "DISPLAY")

    Dim OutMail As Object

    'Criação e chamada do Objeto Outlook
    Set OutMail = Outlook_App.CreateItem(olMailItem)

    Application.DisplayAlerts = False
    With OutMail
        .To = emails("TO")
        .CC = emails("CC")
        .BCC = ""
        .Subject = title
        .HTMLBody = HTML_BODY
        'O trecho abaixo anexa a planilha ao e-mail
        '.Attachments.Add ActiveWorkbook.FullName

    End With

    If (action = "SEND") Then
        OutMail.Send
    ElseIf (action = "DISPLAY") Then
        OutMail.Display
    End If

    Application.DisplayAlerts = True
    'Resetando a sessão
    Set OutMail = Nothing
    Set Outlook_App = Nothing

End Function

