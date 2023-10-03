VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Xdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'namespace=xvba_modules\Xdb



'/*
'
'Class for Data Base Connections
'@author Alberto Eduardo
'@since 14/03/2021
'
'*/



Option Explicit

'Data Base Connection
Public cn As ADODB.Connection
Public source As String
Public dbFolderPath As String
Public dbName As String
Public user As String
Public password As String




Public Function OpenConnection() As ADODB.Connection

  Dim dbPath As String

  ' Handle relative path for the dbPath of Access and SQLite database files
  If (source = "Access" Or source = "SQLite" Or source = "Excel") And Not source Like "?:\*" Then
    dbPath = dbFolderPath & "\" & dbName
  Else
    dbPath = source
  End If

  ' Build the connection string depending on the source
  Dim connectionString As String
  Select Case source
   Case "Access"
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
   Case "Excel"
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Extended Properties=""Excel 8.0;"""
   Case "MySQL"
    connectionString = "Driver={MySQL ODBC 5.2a Driver};Server=" & dbPath & ";Database=test;UID=" & user & ";PWD=" & password
   Case "PostgreSQL"
    connectionString = "Driver={PostgreSQL ANSI};Server=" & dbPath & ";Database=test;UID=" & user & ";PWD=" & password
   Case "SQLite"
    connectionString = "Driver={SQLite3 ODBC Driver};Database=" & dbPath
  End Select

  ' Create and open a new connection to the selected source
  Set OpenConnection = New ADODB.Connection
  Call OpenConnection.Open(connectionString)

End Function




'/*
'
'Sql Insert
'
'@param[String] tb = Table name
'@param[Dictionary] data = data(tab_field_name)
'/*
Public Function Insert(ByVal tb As String, ByVal data As Variant) As Long

  Dim i As Integer
  Dim tableFileds As Variant
  tableFileds = data.Keys
  Dim sql As String
  Dim sqlInit As String
  Dim sqlValues As String
  sqlInit = "INSERT INTO " & tb & " ("
  sqlValues = "VALUES("

  For i = 0 To data.count - 1

    If (data.count = 1) Then
      sqlInit = sqlInit & tableFileds(i) & ") VALUES('" & data(tableFileds(i)) & "')"
    Else
      sqlInit = sqlInit & tableFileds(i) & ","
      sqlValues = sqlValues & "'" & data(tableFileds(i)) & "',"
    End If

    Next

    If (data.count > 1) Then
      sql = Left(sqlInit, Len(sqlInit) - 1) & ") " & Left(sqlValues, Len(sqlValues) - 1) & ")"
    Else
      sql = sqlInit
    End If
Debug.Print sql
    cn.Execute (sql)
    On Error Resume Next
    Insert = GetLastInsertedId(tb)

End Function


Public Function selectFrom(tb As String, Optional ByVal fields As Variant, Optional where As String = "") As ADODB.Recordset

  Dim i As Integer
  Dim sql As String
  Dim fileds_str As String
  sql = "SELECT "
  If (Not fields) Then
    sql = "SELECT * FROM " & tb & where
  Else
    For i = 0 To fields.count - 1
      fileds_str = fileds_str & fields(i) & ","
      Next
      sql = sql & Left(fileds_str, Len(fileds_str) - 1) & " FROM " & tb & where
    End If

    Set selectFrom = cn.Execute(sql)

End Function



Public Function update(tb As String, ByVal fields As Variant, where As String) As ADODB.Recordset

  Dim i As Integer
  Dim sql As String
  Dim fileds_str As String
  Dim fieldNames As Variant
  fieldNames = fields.Keys

  sql = "UPDATE " & tb & "  SET  "

  For i = 0 To fields.count - 1

    fileds_str = fileds_str & fieldNames(i) & " ='" & fields(fieldNames(i)) & "',"

  Next i

  sql = sql & Left(fileds_str, Len(fileds_str) - 1) & " Where " & where
  Debug.Print sql
  Set update = cn.Execute(sql)

End Function



Public Function getAll(ByVal table As String, Optional ByVal searchWord As String = "", Optional ByVal filterType As String = "name") As Variant


  Dim sqlStrQuery As String


  If (searchWord = "") Then
    sqlStrQuery = "SELECT  * FROM  " & table & "   ORDER BY id DESC "
  Else

    sqlStrQuery = "SELECT  * FROM  " & table & "    Where  " & filterType & "   LIKE '%" & searchWord & "%'  ORDER BY id DESC "
  End If

  Set getAll = cn.Execute(sqlStrQuery)


End Function


Public Function SelectX(query_name As String, Optional ByVal params As Variant) As ADODB.Recordset

  Dim i As Integer
  Dim sql As String
  Dim sqlWithParams As String
  Dim fileds_str As String
  Dim replaceField As String
  Dim sql_query As String

  Dim tableFileds As Variant
  tableFileds = params.Keys

  'Get the query from db
  Dim queryResp  As ADODB.Recordset
  sql_query = "SELECT sql_query FROM app_querys WHERE name = '" & query_name & "' LIMIT 1"
  Set queryResp = cn.Execute(sql_query)
  queryResp.MoveFirst
  sql = queryResp!sql_query
  'Check if there are params to replace in query

  sqlWithParams = sql

  For i = 0 To params.count - 1
    replaceField = "[R" & i & "]"
    sqlWithParams = Replace(sqlWithParams, replaceField, params(tableFileds(i)))
  Next i

Debug.Print sqlWithParams

  Set SelectX = cn.Execute(sqlWithParams)

End Function

'/*
'
'Get last Inserted ID Bay table
'
'@param <String>  tableNAme
'
'@return <Long>  Id
'*/
Public Function GetLastInsertedId(tableName As String) As Long

  Dim rs As ADODB.Recordset
  Dim sqlLastInsertedIdQuery As String

  sqlLastInsertedIdQuery = "SELECT seq FROM sqlite_sequence WHERE name='" & tableName & "'"

  'Get Last inserted Id
  Set rs = cn.Execute(sqlLastInsertedIdQuery)
  GetLastInsertedId = rs.fields("seq").Value

End Function



'/*
'
'Delete  By ID
'
'*/
Public Function deleteById(ByVal idData As String, ByVal tb As String)


  Dim sqlStrQuery As String

  sqlStrQuery = " DELETE   FROM  " & tb & "  Where id = " & idData

  On Error Resume Next
  cn.Execute (sqlStrQuery)



End Function
