Attribute VB_Name = "db_doc_review"


'namespace=vba-files\DataBase\Documents


' Constants For database connection
Private Const tableName As String = "documents_reviews"

' Array of valid table fields
Private Const field1 As String = "rev_code"
Private Const field2 As String = "issue"
Private Const Field3 As String = "status"
Private Const Field4 As String = "status_date"
Private Const Field5 As String = "file_path"
Private Const Field6 As String = "file_name"
Private Const Field7 As String = "file_extension"
Private Const Field8 As String = "next_review"
Private Const Field9 As String = "next_issue"
Private Const Field10 As String = "request_doc_id"

' Add more field constants As needed

'/*
'
'
'data as Arrray
'*/
Public Function update(data As Variant)

    Dim tableFields As Variant
    tableFields = Array(field1, field2, Field3, Field4, Field5, Field6, Field7, Field8, Field9, Field10) ' Replace With your actual table field names



    Dim condition As String
    condition = "ID = 123" ' Replace With your actual condition

    Dim success As Boolean
    success = UpdateTable(data, condition, tableFields)

    If success Then
        MsgBox "Update successful!"
    Else
        MsgBox "Update failed!"
    End If

End Function

Private Function UpdateTable(ByVal data As Variant, ByVal condition As String, ByVal tableFields As Variant) As Boolean
    On Error GoTo ErrorHandler

        Dim i As Integer
        Dim fieldNames As String
        Dim fieldParams As String

        fieldNames = ""
        fieldParams = ""

        For i = 0 To UBound(data)
            If Not IsFieldValid(data(i), tableFields) Then
                MsgBox "Invalid field name: " & data(i)
             Exit Function
            End If

            fieldNames = fieldNames & data(i) & " = ?, "
            fieldParams = fieldParams & "param" & i & ","
        Next i

        fieldNames = Left(fieldNames, Len(fieldNames) - 2) ' Remove the trailing comma And space
        fieldParams = Left(fieldParams, Len(fieldParams) - 1) ' Remove the trailing comma

        Dim sqlStrQuery As String
        sqlStrQuery = "UPDATE " & tableName & " Set " & fieldNames & " WHERE " & condition

        Dim cmd As Object
        Set cmd = database.CreateCommand()

        database.BeginTrans ' Start the transaction

        For i = 0 To UBound(data)
            cmd.Parameters.Append cmd.CreateParameter("param" & i, adVarChar, adParamInput, Len(data(i)), data(i))
            ' Modify adVarChar To match the appropriate data type of the fields
        Next i

        cmd.CommandText = sqlStrQuery
        cmd.Execute

        database.CommitTrans ' Commit the transaction If successful

        UpdateTable = True ' Success

     Exit Function

ErrorHandler:
        database.RollbackTrans ' Rollback the transaction If an error occurs
        ' Handle the error
        UpdateTable = False ' Error occurred
End Function

Public Function IsFieldValid(ByVal field As String, ByVal tableFields As Variant) As Boolean
    ' Check If the provided field is valid
    Dim i As Integer
    For i = 0 To UBound(tableFields)
        If field = tableFields(i) Then
            IsFieldValid = True
         Exit Function
        End If
    Next i
    IsFieldValid = False
End Function
