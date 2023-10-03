Attribute VB_Name = "DateHelpers"


'namespace=vba-files\Helpers




'/*
'
'
'Convert Date from SQLite  for BR format
'
'*/
Public Function FormatSQliteToDate(ByVal db_date As String) As Variant

If (db_date <> "" And db_date <> "0") Then
    On Error GoTo error_handler
    FormatSQliteToDate = format(CDate(db_date), "DD/MM/YYYY")


Else
FormatSQliteToDate = ""
End If

    Exit Function
error_handler:
    Call Alert.Show("Erro ao converter a Data", "", 2500)
End Function


'/*
'
'
'Convert Date to SQLite format
'
'*/
Public Function FormatDateToSQlite(ByVal db_date As String) As String
If (db_date <> "" And db_date <> "0") Then
    On Error GoTo error_handler
    FormatDateToSQlite = format(CDate(db_date), "YYYY-MM-DD")
Else
FormatDateToSQlite = ""
End If
        Exit Function
error_handler:
        Call Alert.Show("Erro ao converter a Data", "", 2500)

End Function
