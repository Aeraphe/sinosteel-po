Attribute VB_Name = "XdbFactory"

'namespace=xvba_modules\Xdb



Function Create(Optional DataBaseName As String = "database.db", _
    Optional bdType As String = "SQLite", _
    Optional dbPath As String, _
    Optional userName As String, _
    Optional pass As String) As Object

    If (dbPath = "") Then
        dbPath = config_sheet.Range("CONFIG_DATABASE_PATH").Value
    End If

    Dim database  As Object
    Set database = New Xdb

    database.source = bdType
    database.dbName = DataBaseName
    database.dbFolderPath = dbPath
    database.user = userName
    database.password = pass

    Set database.cn = database.OpenConnection()

    database.cn.Execute ("PRAGMA foreign_keys = ON")
    Set Create = database

End Function


'/*
'
'Check If the response from recorset is Not null And return the value
'If the response is null return ""
'
'*/
Public Function getData(ByRef rs As ADODB.Recordset, ByVal field As String, Optional respType As String = "STRING") As Variant

    Dim response As Variant
    On Error Resume Next
    response = rs.fields(field).Value

    If (Not IsNull(response)) Then

        If (respType = "STRING") Then
            getData = response
        End If

        If (respType = "LONG") Then
            getData = CLng(response)
        End If

        If (respType = "INT") Then
            getData = CInt(response)
        End If

        If (respType = "SQL_DATE") Then
            getData = format(CDate(db_date), "DD/MM/YYYY")
        End If

    ElseIf (respType = "STRING") Then
        getData = ""
    ElseIf (respType = "LONG") Then
        getData = 0
    ElseIf (respType = "SQL_DATE") Then
        getData = ""
    Else
        getData = "0"

    End If

End Function



Public Function SelectX(ByVal query As String, ByVal data As Object) As Recordset

    Dim database As Object
    Set database = XdbFactory.Create

    Set SelectX = database.SelectX(query, data)

End Function



Public Function ExecuteSQLTransaction(sqlStrQuery As String) As Boolean

    On Error GoTo ErrorHandler

        Dim database As Object
        Dim statements() As String
        Dim i As Long

        Set database = XdbFactory.Create

        ' Split the SQL statements
        statements = Split(sqlStrQuery, ";")

        database.cn.Execute "BEGIN TRANSACTION;"

        ' Execute each statement one-by-one
        For i = LBound(statements) To UBound(statements) - 1 ' -1 because the last item after split is empty
            If Trim(statements(i)) <> "" Then ' Check To ensure the statement is Not just white space
                database.cn.Execute Trim(statements(i))
            End If
        Next i

        database.cn.Execute "COMMIT;"

        ExecuteSQLTransaction = True

     Exit Function

ErrorHandler:
        database.cn.Execute "ROLLBACK;"
        MsgBox "Error " & Err.Number & ": " & Err.description
        ExecuteSQLTransaction = False

End Function
