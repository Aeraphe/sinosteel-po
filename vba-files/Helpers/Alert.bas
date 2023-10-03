Attribute VB_Name = "Alert"


'namespace=vba-files\Helpers


'/*
'
'
'
'*/
Public Function Show(title As String, description As String, delay As Long)

  UserFormAlert.Show
  UserFormAlert.Label1.Caption = title
  UserFormAlert.labelInfo.Caption = description
  UserFormAlert.Repaint
  Xhelper.waitMs (delay)
  UserFormAlert.Hide

End Function
