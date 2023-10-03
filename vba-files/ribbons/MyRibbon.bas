Attribute VB_Name = "MyRibbon"

'namespace=vba-files/ribbons




Public Sub showAddPurchaseOrder(ByRef control As Office.IRibbonControl)
    On Error Resume Next
    purchase_order_add.Show
End Sub




'/*
'
'Version Form Show
'
'
'*/
Public Sub showVersionForm(ByRef control As Office.IRibbonControl)

    VersionForm.titleLabel.Caption = "Controle de Documentos"
    VersionForm.subTitleLabel.Caption = "Ver: 1.91.0beta (luci)"
    VersionForm.Repaint
    VersionForm.Show

End Sub
'/*
'
'
'
'
'*/
Public Sub showLoggingForm(ByRef control As Office.IRibbonControl)
    On Error Resume Next
    login_form.Show
End Sub

'
'
'
'*/
Public Sub showAddUserForm(ByRef control As Office.IRibbonControl)
On Error Resume Next
    user_add_form.Show
End Sub

'/*
'
'
'
'
'*/
Public Sub showUserInfoForm(ByRef control As Office.IRibbonControl)
On Error Resume Next
    user_change_info_form.Show
End Sub



'/*
'
'Loggout
'
'
'*/
Public Sub showLogoutForm(ByRef control As Office.IRibbonControl)

    Dim answer As Integer

    answer = MsgBox("Quer Sair do Sistema?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert!!!")

    If (answer = vbYes) Then
        auth.loggout
        core_user_config.UpdateStayLoggedInValue False
    End If
End Sub

