Attribute VB_Name = "service_email"

'namespace=vba-files\Services




'/*
'
'Get all emails to send message
'
'*/
 Function get_emails(ByVal recipient_id As String) As Object

    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    Dim emails_to As String
    Dim emails_cc As String

    emails_cc = ""
    emails_to = ""


    data("TO") = emails_to
    data("CC") = emails_cc

    Set get_emails = data
End Function




'/*
'
'Semd email message
'
'*/
Function Send(emails As Object, ByVal title As String, ByVal HTML_BODY As String, ByRef Outlook_App As Object, Optional ByRef attachs As Object, Optional action As String = "DISPLAY")
    Dim OutMail As Object

    ' Create an Outlook mail item object
    Set OutMail = Outlook_App.CreateItem(olMailItem)

    ' Disable alerts
    Application.DisplayAlerts = False

    With OutMail
        ' Set email recipient, CC, BCC, subject, and body
        .To = emails("TO")
        .CC = emails("CC")
        .BCC = ""
        .Subject = title
        .HTMLBody = HTML_BODY

        ' Add attachments if attachs parameter is not missing
        If Not IsMissing(attachs) And Not attachs Is Nothing Then
            For Each varKey In attachs.Keys()
            If (Not IsEmpty(varKey)) Then
                .Attachments.Add attachs(varKey)
                End If
            Next
        End If

        ' Send or display the email based on the action parameter
        If action = "SEND" Then
            .Send
        ElseIf action = "DISPLAY" Then
            .Display
        End If
    End With

    ' Enable alerts
    Application.DisplayAlerts = True

    ' Reset the objects
    Set OutMail = Nothing
    Set Outlook_App = Nothing
End Function
