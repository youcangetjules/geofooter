Attribute VB_Name = "Module3"
Option Explicit

Private Const adVarWChar As Long = 202
Private Const adParamInput As Long = 1

' Returns the GURI for the given parameters, or "" if not found
Public Function GetGURI(sender As String, recipients As String, Subject As String, dt_str As String, avg_risk As String) As String
    Dim conn As Object, rs As Object, cmd As Object
    Dim dbPath As String
    dbPath = "C:\GeoFooter\guri_records.db" ' Adjust path if needed

    On Error GoTo ErrHandler

    Dim cs As String
    cs = GetSQLiteConnectionString(dbPath)
    If cs = "" Then
        Module1.WriteLog "GetGURI: SQLite ODBC driver not available, returning empty string"
        GetGURI = ""
        Exit Function
    End If

    Set conn = CreateObject("ADODB.Connection")
    conn.Open cs

    Dim sql As String
    sql = "SELECT guri FROM guri_records WHERE sender=? AND recipients=? AND subject=? AND [datetime]=? AND avg_risk=?"

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, sender)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, recipients)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, Subject)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, dt_str)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, avg_risk)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        GetGURI = rs.Fields("guri").Value
    Else
        GetGURI = ""
    End If

    rs.Close
    conn.Close
    Exit Function

ErrHandler:
    Module1.WriteLog "GetGURI: Error #" & Err.Number & " - " & Err.Description
    GetGURI = ""
    If Not rs Is Nothing Then On Error Resume Next: rs.Close
    If Not conn Is Nothing Then On Error Resume Next: conn.Close
End Function

' Insert a new GURI record (optional)
Public Sub InsertGURI(guri As String, sender As String, recipients As String, Subject As String, dt_str As String, avg_risk As String, random_block As String)
    Dim conn As Object, cmd As Object
    Dim dbPath As String
    dbPath = "C:\GeoFooter\guri_records.db" ' Adjust path if needed

    On Error GoTo ErrHandler

    Dim cs As String
    cs = GetSQLiteConnectionString(dbPath)
    If cs = "" Then
        Module1.WriteLog "InsertGURI: SQLite ODBC driver not available, skipping database insert"
        Exit Sub
    End If

    Set conn = CreateObject("ADODB.Connection")
    conn.Open cs

    Dim sql As String
    sql = "INSERT OR IGNORE INTO guri_records (guri, sender, recipients, subject, [datetime], avg_risk, random_block) VALUES (?, ?, ?, ?, ?, ?, ?)"

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, guri)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, sender)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, recipients)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, Subject)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, dt_str)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, avg_risk)
    cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 255, random_block)

    cmd.Execute

    conn.Close
    Exit Sub

ErrHandler:
    Module1.WriteLog "InsertGURI: Error #" & Err.Number & " - " & Err.Description
    If Not conn Is Nothing Then On Error Resume Next: conn.Close
End Sub

' Try common SQLite ODBC driver names and return a working connection string, or "" if none found
Private Function GetSQLiteConnectionString(dbPath As String) As String
    Dim drivers(1 To 3) As String
    drivers(1) = "SQLite3 ODBC Driver"
    drivers(2) = "SQLite ODBC Driver"
    drivers(3) = "SQLite3 Driver"

    Dim i As Long
    For i = LBound(drivers) To UBound(drivers)
        On Error Resume Next
        Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
        cn.Open "Driver={" & drivers(i) & "};Database=" & dbPath & ";"
        If Err.Number = 0 Then
            cn.Close
            GetSQLiteConnectionString = "Driver={" & drivers(i) & "};Database=" & dbPath & ";"
            Exit Function
        End If
        Err.Clear
    Next i
    GetSQLiteConnectionString = "" ' not found
End Function
