Attribute VB_Name = "MSCANSettings"
'===============================================================================
' MSCANSettings - AES account scan settings (no UserForm required)
' Persists which Outlook mailboxes get inbox watchers / automatic footers.
'===============================================================================
Option Explicit

Private Const SETTINGS_FILE_NAME As String = "aes_scan_accounts.ini"
Private Const ROUTE_RISK_FILE_NAME As String = "aes_risk_route.json"
Private Const LOGGING_FILE_NAME As String = "aes_logging.json"

Private mSettingsLoaded As Boolean
Private mAccountStates As Object     ' Scripting.Dictionary: storeId -> enabled (Boolean)
Private mAccountResponses As Object  ' Scripting.Dictionary: storeId -> scan responses (Boolean)
Private mAccountInCc As Object       ' Scripting.Dictionary: storeId -> scan In Cc (Boolean)

Public Sub EnsureSettingsLoaded()
    If mSettingsLoaded Then Exit Sub
    Set mAccountStates = CreateObject("Scripting.Dictionary")
    mAccountStates.CompareMode = vbTextCompare
    Set mAccountResponses = CreateObject("Scripting.Dictionary")
    mAccountResponses.CompareMode = vbTextCompare
    Set mAccountInCc = CreateObject("Scripting.Dictionary")
    mAccountInCc.CompareMode = vbTextCompare
    LoadSettingsFromFile
    mSettingsLoaded = True
End Sub

Public Function GetSettingsFilePath() As String
    Dim baseDir As String
    baseDir = ResolveSettingsDirectory()
    If Len(baseDir) = 0 Then Exit Function
    GetSettingsFilePath = baseDir & "\" & SETTINGS_FILE_NAME
End Function

Public Function GetRouteRiskFilePath() As String
    Dim baseDir As String
    baseDir = ResolveSettingsDirectory()
    If Len(baseDir) = 0 Then Exit Function
    GetRouteRiskFilePath = baseDir & "\" & ROUTE_RISK_FILE_NAME
End Function

Public Function GetLoggingSettingsFilePath() As String
    Dim baseDir As String
    baseDir = ResolveSettingsDirectory()
    If Len(baseDir) = 0 Then Exit Function
    GetLoggingSettingsFilePath = baseDir & "\" & LOGGING_FILE_NAME
End Function

Private Function ResolveSettingsDirectory() As String
    Dim paths As Variant
    Dim p As Variant
    Dim fso As Object

    paths = Array( _
        Environ$("LOCALAPPDATA") & "\GeoFooter", _
        Environ$("TEMP") & "\GeoFooter", _
        MSCANPaths.GetInstallRoot())

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each p In paths
        On Error Resume Next
        If Not fso.FolderExists(CStr(p)) Then fso.CreateFolder CStr(p)
        If fso.FolderExists(CStr(p)) Then
            ResolveSettingsDirectory = CStr(p)
            Exit Function
        End If
        On Error GoTo 0
    Next p
End Function

Public Function IsAccountScanEnabled(ByVal acc As Outlook.Account) As Boolean
    On Error GoTo EH

    If acc Is Nothing Then
        IsAccountScanEnabled = False
        Exit Function
    End If

    EnsureSettingsLoaded

    Dim storeId As String
    storeId = AccountStoreId(acc)
    If Len(storeId) = 0 Then
        IsAccountScanEnabled = True
        Exit Function
    End If

    If mAccountStates.Exists(storeId) Then
        IsAccountScanEnabled = CBool(mAccountStates(storeId))
    Else
        IsAccountScanEnabled = True
    End If
    Exit Function

EH:
    IsAccountScanEnabled = True
End Function

Public Function IsMailItemAccountScanEnabled(ByVal mail As Object) As Boolean
    On Error GoTo EH

    If mail Is Nothing Then
        IsMailItemAccountScanEnabled = False
        Exit Function
    End If

    Dim storeId As String
    storeId = mail.Parent.StoreID

    Dim acc As Outlook.Account
    Dim i As Long
    For i = 1 To Application.Session.Accounts.Count
        Set acc = Application.Session.Accounts.Item(i)
        If StrComp(AccountStoreId(acc), storeId, vbBinaryCompare) = 0 Then
            IsMailItemAccountScanEnabled = IsAccountScanEnabled(acc)
            Exit Function
        End If
    Next i

    IsMailItemAccountScanEnabled = True
    Exit Function

EH:
    IsMailItemAccountScanEnabled = True
End Function

Public Sub SetAccountScanEnabled(ByVal acc As Outlook.Account, ByVal enabled As Boolean)
    EnsureSettingsLoaded
    Dim storeId As String
    storeId = AccountStoreId(acc)
    If Len(storeId) = 0 Then Exit Sub
    mAccountStates(storeId) = enabled
End Sub

Public Function IsAccountResponsesEnabled(ByVal acc As Outlook.Account) As Boolean
    On Error GoTo EH
    If acc Is Nothing Then
        IsAccountResponsesEnabled = False
        Exit Function
    End If
    EnsureSettingsLoaded
    Dim storeId As String
    storeId = AccountStoreId(acc)
    If Len(storeId) = 0 Then
        IsAccountResponsesEnabled = True
        Exit Function
    End If
    If mAccountResponses.Exists(storeId) Then
        IsAccountResponsesEnabled = CBool(mAccountResponses(storeId))
    Else
        IsAccountResponsesEnabled = True
    End If
    Exit Function
EH:
    IsAccountResponsesEnabled = True
End Function

Public Function IsAccountInCcEnabled(ByVal acc As Outlook.Account) As Boolean
    On Error GoTo EH
    If acc Is Nothing Then
        IsAccountInCcEnabled = False
        Exit Function
    End If
    EnsureSettingsLoaded
    Dim storeId As String
    storeId = AccountStoreId(acc)
    If Len(storeId) = 0 Then
        IsAccountInCcEnabled = True
        Exit Function
    End If
    If mAccountInCc.Exists(storeId) Then
        IsAccountInCcEnabled = CBool(mAccountInCc(storeId))
    Else
        IsAccountInCcEnabled = True
    End If
    Exit Function
EH:
    IsAccountInCcEnabled = True
End Function

Public Sub SetAccountResponsesEnabled(ByVal acc As Outlook.Account, ByVal enabled As Boolean)
    EnsureSettingsLoaded
    Dim storeId As String
    storeId = AccountStoreId(acc)
    If Len(storeId) = 0 Then Exit Sub
    mAccountResponses(storeId) = enabled
End Sub

Public Sub SetAccountInCcEnabled(ByVal acc As Outlook.Account, ByVal enabled As Boolean)
    EnsureSettingsLoaded
    Dim storeId As String
    storeId = AccountStoreId(acc)
    If Len(storeId) = 0 Then Exit Sub
    mAccountInCc(storeId) = enabled
End Sub

Public Sub ShowSettingsDialog()
    On Error GoTo EH

    EnsureSettingsLoaded
    SyncAccountsFromOutlook

    Dim cancelled As Boolean
    cancelled = True

    If TryShowPythonSettingsDialog(cancelled) Then
        ' Python dialog handled UI; cancelled already set
    Else
        cancelled = ShowInputBoxSettingsDialog()
    End If

    If cancelled Then
        mSettingsLoaded = False
        Set mAccountStates = Nothing
        Set mAccountResponses = Nothing
        Set mAccountInCc = Nothing
        EnsureSettingsLoaded
        MSCANModStatus.ShowStatus "AES settings unchanged."
        Exit Sub
    End If

    SaveSettingsToFile
    MSCANModLogging.ReloadLoggingSettings
    MSCANModLogging.WriteLogAudit "MSCANSettings: Saved account scan settings."
    If Len(Dir$(GetRouteRiskFilePath())) > 0 Then
        MSCANModLogging.WriteLogAudit "MSCANSettings: Route/ASN risk config at " & GetRouteRiskFilePath()
    End If
    If Len(Dir$(GetLoggingSettingsFilePath())) > 0 Then
        MSCANModLogging.WriteLogAudit "MSCANSettings: Logging config at " & GetLoggingSettingsFilePath()
    End If

    If MSCANModWatchers.IsServiceEnabled() Then
        MSCANModWatchers.CheckWatchers
    End If

    MSCANModStatus.ShowStatus "AES settings saved. Active watchers: " & MSCANModWatchers.GetWatcherCount()
    Exit Sub

EH:
    MsgBox "AES settings error: " & Err.Description, vbExclamation, "AES Settings"
End Sub

' Returns True if Python dialog ran (success or cancel). False = use InputBox fallback.
' cancelled=True means user cancelled / no save.
Private Function TryShowPythonSettingsDialog(ByRef cancelled As Boolean) As Boolean
    On Error GoTo EH

    TryShowPythonSettingsDialog = False
    cancelled = True

    Dim pythonExe As String
    Dim scriptPath As String
    Dim accountsPath As String
    Dim outPath As String

    pythonExe = ResolveSettingsPythonExe()
    scriptPath = ResolveSettingsDialogScript()
    If Len(pythonExe) = 0 Or Len(scriptPath) = 0 Then
        MSCANModLogging.WriteLog "MSCANSettings: Python settings dialog unavailable; InputBox fallback."
        Exit Function
    End If

    accountsPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_settings_accounts.json"
    outPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_settings_result.json"
    Dim routePath As String
    routePath = GetRouteRiskFilePath()
    If Len(routePath) = 0 Then routePath = Environ$("LOCALAPPDATA") & "\GeoFooter\" & ROUTE_RISK_FILE_NAME
    Dim loggingPath As String
    loggingPath = GetLoggingSettingsFilePath()
    If Len(loggingPath) = 0 Then loggingPath = Environ$("LOCALAPPDATA") & "\GeoFooter\" & LOGGING_FILE_NAME

    If Not WriteAccountsJson(accountsPath) Then
        MSCANModLogging.WriteLog "MSCANSettings: could not write accounts JSON; InputBox fallback."
        Exit Function
    End If

    On Error Resume Next
    Kill outPath
    Err.Clear
    On Error GoTo EH

    Dim cmd As String
    cmd = """" & pythonExe & """ """ & scriptPath & """" & _
          " --accounts " & ShellQuote(accountsPath) & _
          " --out " & ShellQuote(outPath) & _
          " --route " & ShellQuote(routePath) & _
          " --logging " & ShellQuote(loggingPath)

    Dim shell As Object
    Dim exitCode As Long
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(cmd, 1, True)

    If exitCode <> 0 Then
        MSCANModLogging.WriteLog "MSCANSettings: Python dialog exit " & exitCode & "; InputBox fallback."
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(outPath) Then
        MSCANModLogging.WriteLog "MSCANSettings: result JSON missing; InputBox fallback."
        Exit Function
    End If

    Dim raw As String
    raw = ReadAllTextFile(outPath)
    If Len(raw) = 0 Then
        MSCANModLogging.WriteLog "MSCANSettings: result JSON empty; InputBox fallback."
        Exit Function
    End If

    TryShowPythonSettingsDialog = True

    If JsonBoolTrue(raw, "cancelled") Then
        cancelled = True
        Exit Function
    End If

    If Not ApplySettingsFromResultJson(raw) Then
        cancelled = True
        MSCANModLogging.WriteLog "MSCANSettings: could not parse result JSON; treating as cancel."
        Exit Function
    End If

    cancelled = False
    Exit Function

EH:
    MSCANModLogging.WriteLog "MSCANSettings.TryShowPythonSettingsDialog error: #" & Err.Number & " - " & Err.Description
    TryShowPythonSettingsDialog = False
    cancelled = True
End Function

' Returns True if user cancelled (no save).
Private Function ShowInputBoxSettingsDialog() As Boolean
    On Error GoTo EH

    Dim cancelled As Boolean
    cancelled = False

    Do
        Dim prompt As String
        prompt = BuildSettingsPrompt()
        Dim choice As String
        choice = Trim$(InputBox(prompt, "AES Settings - Scan Accounts", ""))

        If Len(choice) = 0 Then
            cancelled = True
            Exit Do
        End If

        Select Case UCase$(choice)
            Case "C", "CANCEL"
                cancelled = True
                Exit Do
            Case "S", "SAVE"
                Exit Do
            Case "A", "ALL"
                SetAllAccounts True
            Case "N", "NONE"
                SetAllAccounts False
            Case Else
                If Not IsNumeric(choice) Then
                    MsgBox "Enter an account number, A, N, S, or C.", vbExclamation, "AES Settings"
                Else
                    ToggleAccountByIndex CLng(choice)
                End If
        End Select
    Loop

    ShowInputBoxSettingsDialog = cancelled
    Exit Function

EH:
    ShowInputBoxSettingsDialog = True
End Function

Private Function WriteAccountsJson(ByVal filePath As String) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Dim ts As Object
    Dim acc As Outlook.Account
    Dim i As Long
    Dim first As Boolean
    Dim storeId As String
    Dim displayName As String
    Dim smtp As String
    Dim enabledFlag As String
    Dim responsesFlag As String
    Dim inCcFlag As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureParentFolder filePath
    Set ts = fso.CreateTextFile(filePath, True, False)
    ts.WriteLine "{"
    ts.WriteLine "  ""accounts"": ["
    first = True

    For i = 1 To Application.Session.Accounts.Count
        Set acc = Application.Session.Accounts.Item(i)
        storeId = AccountStoreId(acc)
        If Len(storeId) = 0 Then GoTo NextAcc
        displayName = JsonEscape(acc.DisplayName)
        smtp = JsonEscape(AccountSmtpAddress(acc))
        If IsAccountScanEnabled(acc) Then
            enabledFlag = "true"
        Else
            enabledFlag = "false"
        End If
        If IsAccountResponsesEnabled(acc) Then
            responsesFlag = "true"
        Else
            responsesFlag = "false"
        End If
        If IsAccountInCcEnabled(acc) Then
            inCcFlag = "true"
        Else
            inCcFlag = "false"
        End If
        If Not first Then ts.WriteLine ","
        first = False
        ts.Write "    {""store_id"":""" & JsonEscape(storeId) & """,""display"":""" & displayName & _
                 """,""smtp"":""" & smtp & """,""enabled"":" & enabledFlag & _
                 ",""responses"":" & responsesFlag & ",""in_cc"":" & inCcFlag & "}"
NextAcc:
    Next i

    ts.WriteLine ""
    ts.WriteLine "  ],"
    ts.WriteLine "  ""route_risk"": " & RouteRiskJsonObject() & ","
    ts.WriteLine "  ""logging"": " & LoggingJsonObject()
    ts.WriteLine "}"
    ts.Close
    WriteAccountsJson = True
    Exit Function

EH:
    WriteAccountsJson = False
End Function

' Compact JSON object for route/ASN risk (embedded in accounts JSON).
Private Function RouteRiskJsonObject() As String
    On Error Resume Next

    Dim path As String
    Dim raw As String
    path = GetRouteRiskFilePath()
    If Len(path) > 0 Then
        raw = Trim$(ReadAllTextFile(path))
        If Len(raw) > 0 And Left$(raw, 1) = "{" Then
            RouteRiskJsonObject = raw
            Exit Function
        End If
    End If

    RouteRiskJsonObject = "{""provider"":""abuseipdb"",""abuse_threshold"":25," & _
                          """extra_high_risk_asns"":[],""enabled"":true}"
End Function

Private Function LoggingJsonObject() As String
    On Error Resume Next

    Dim path As String
    Dim raw As String
    path = GetLoggingSettingsFilePath()
    If Len(path) > 0 Then
        raw = Trim$(ReadAllTextFile(path))
        If Len(raw) > 0 And Left$(raw, 1) = "{" Then
            LoggingJsonObject = raw
            Exit Function
        End If
    End If

    LoggingJsonObject = "{""enabled"":true,""levels"":{""info"":true,""audit"":true," & _
                        """warn"":true,""debug"":false}}"
End Function

Private Function ApplySettingsFromResultJson(ByVal raw As String) As Boolean
    On Error GoTo EH

    Dim acc As Outlook.Account
    Dim storeId As String
    Dim i As Long
    Dim applied As Long

    applied = 0
    For i = 1 To Application.Session.Accounts.Count
        Set acc = Application.Session.Accounts.Item(i)
        storeId = AccountStoreId(acc)
        If Len(storeId) = 0 Then GoTo NextApply

        Dim enabled As Boolean
        Dim responses As Boolean
        Dim inCc As Boolean
        If TryGetBoolForStoreId(raw, storeId, "enabled", enabled) Then
            SetAccountScanEnabled acc, enabled
            applied = applied + 1
        End If
        If TryGetBoolForStoreId(raw, storeId, "responses", responses) Then
            SetAccountResponsesEnabled acc, responses
        Else
            SetAccountResponsesEnabled acc, True
        End If
        If TryGetBoolForStoreId(raw, storeId, "in_cc", inCc) Then
            SetAccountInCcEnabled acc, inCc
        Else
            SetAccountInCcEnabled acc, True
        End If
NextApply:
    Next i

    ApplySettingsFromResultJson = (applied > 0)
    Exit Function

EH:
    ApplySettingsFromResultJson = False
End Function

Private Function TryGetBoolForStoreId( _
    ByVal raw As String, _
    ByVal storeId As String, _
    ByVal keyName As String, _
    ByRef value As Boolean) As Boolean

    On Error Resume Next

    Dim needle As String
    Dim pos As Long
    Dim chunk As String
    Dim keyPos As Long
    Dim escapedId As String
    Dim endBrace As Long

    escapedId = JsonEscape(storeId)
    needle = """store_id"":""" & escapedId & """"
    pos = InStr(1, raw, needle, vbBinaryCompare)
    If pos = 0 Then
        needle = """store_id"": """ & escapedId & """"
        pos = InStr(1, raw, needle, vbBinaryCompare)
    End If
    If pos = 0 Then
        TryGetBoolForStoreId = False
        Exit Function
    End If

    ' Account objects are flat, so the entry ends at the next "}". Never use a
    ' fixed-size window: Outlook store IDs alone can exceed 500 characters,
    ' which used to push "enabled" out of the chunk and void every save.
    endBrace = InStr(pos, raw, "}", vbBinaryCompare)
    If endBrace = 0 Then endBrace = Len(raw)
    chunk = Mid$(raw, pos, endBrace - pos + 1)

    keyPos = InStr(1, chunk, """" & keyName & """", vbTextCompare)
    If keyPos = 0 Then
        TryGetBoolForStoreId = False
        Exit Function
    End If

    chunk = LCase$(Mid$(chunk, keyPos, 40))
    If InStr(1, chunk, "true", vbTextCompare) > 0 Then
        value = True
        TryGetBoolForStoreId = True
    ElseIf InStr(1, chunk, "false", vbTextCompare) > 0 Then
        value = False
        TryGetBoolForStoreId = True
    Else
        TryGetBoolForStoreId = False
    End If
End Function

Private Function JsonBoolTrue(ByVal raw As String, ByVal key As String) As Boolean
    On Error Resume Next
    Dim pattern As String
    Dim startPos As Long
    Dim chunk As String
    pattern = """" & key & """"
    startPos = InStr(1, raw, pattern, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos, raw, ":", vbBinaryCompare)
    If startPos = 0 Then Exit Function
    chunk = LCase$(Mid$(raw, startPos + 1, 12))
    JsonBoolTrue = (InStr(1, chunk, "true", vbTextCompare) > 0)
End Function

Private Function JsonEscape(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "\r")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    JsonEscape = t
End Function

Private Function ResolveSettingsPythonExe() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim paths As Variant
    paths = Array( _
        Environ$("USERPROFILE") & "\AppData\Local\Programs\Python\Python313\pythonw.exe", _
        Environ$("LOCALAPPDATA") & "\Programs\Python\Python313\pythonw.exe", _
        MSCANPaths.GetVenvPythonw(), _
        MSCANPaths.GetVenvPython())

    Dim p As Variant
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            ResolveSettingsPythonExe = CStr(p)
            Exit Function
        End If
    Next p
    ResolveSettingsPythonExe = ""
End Function

Private Function ResolveSettingsDialogScript() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim paths As Variant
    paths = Array( _
        MSCANPaths.GetAesScript("settings_dialog.py"))

    Dim p As Variant
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            ResolveSettingsDialogScript = CStr(p)
            Exit Function
        End If
    Next p
    ResolveSettingsDialogScript = ""
End Function

Private Sub EnsureParentFolder(ByVal filePath As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parent As String
    parent = fso.GetParentFolderName(filePath)
    If Len(parent) = 0 Then Exit Sub
    If Not fso.FolderExists(parent) Then fso.CreateFolder parent
End Sub

Private Function ShellQuote(ByVal value As String) As String
    ShellQuote = """" & Replace(value, """", """""") & """"
End Function

Private Function ReadAllTextFile(ByVal filePath As String) As String
    On Error GoTo EH
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -2)
    If Not ts.AtEndOfStream Then
        ReadAllTextFile = ts.ReadAll
    Else
        ReadAllTextFile = ""
    End If
    ts.Close
    Exit Function
EH:
    ReadAllTextFile = ""
End Function

Public Function GetAccountScanDiagnostics() As String
    Dim info As String
    Dim acc As Outlook.Account
    Dim i As Long
    Dim enabledCount As Long
    Dim totalCount As Long

    EnsureSettingsLoaded
    SyncAccountsFromOutlook

    info = "--- AES ACCOUNT SCAN SETTINGS ---" & vbCrLf
    info = info & "Settings file: " & GetSettingsFilePath() & vbCrLf

    For i = 1 To Application.Session.Accounts.Count
        Set acc = Application.Session.Accounts.Item(i)
        totalCount = totalCount + 1
        Dim enabled As Boolean
        enabled = IsAccountScanEnabled(acc)
        If enabled Then enabledCount = enabledCount + 1
        info = info & "  " & i & ". "
        If enabled Then
            info = info & "[SCAN] "
        Else
            info = info & "[SKIP] "
        End If
        info = info & acc.DisplayName & " (" & AccountSmtpAddress(acc) & ")"
        info = info & "  Responses=" & IIf(IsAccountResponsesEnabled(acc), "ON", "OFF")
        info = info & "  InCc=" & IIf(IsAccountInCcEnabled(acc), "ON", "OFF") & vbCrLf
    Next i

    info = info & "Scan enabled: " & enabledCount & " / " & totalCount & vbCrLf
    info = info & "Active watchers: " & MSCANModWatchers.GetWatcherCount() & vbCrLf
    info = info & vbCrLf
    GetAccountScanDiagnostics = info
End Function

Private Sub SyncAccountsFromOutlook()
    Dim acc As Outlook.Account
    Dim storeId As String

    For Each acc In Application.Session.Accounts
        storeId = AccountStoreId(acc)
        If Len(storeId) > 0 Then
            If Not mAccountStates.Exists(storeId) Then
                mAccountStates(storeId) = True
            End If
            If Not mAccountResponses.Exists(storeId) Then
                mAccountResponses(storeId) = True
            End If
            If Not mAccountInCc.Exists(storeId) Then
                mAccountInCc(storeId) = True
            End If
        End If
    Next acc
End Sub

Private Sub SetAllAccounts(ByVal enabled As Boolean)
    Dim acc As Outlook.Account
    For Each acc In Application.Session.Accounts
        SetAccountScanEnabled acc, enabled
    Next acc
End Sub

Private Sub ToggleAccountByIndex(ByVal index As Long)
    If index < 1 Or index > Application.Session.Accounts.Count Then
        MsgBox "Invalid account number.", vbExclamation, "AES Settings"
        Exit Sub
    End If

    Dim acc As Outlook.Account
    Set acc = Application.Session.Accounts.Item(index)
    SetAccountScanEnabled acc, Not IsAccountScanEnabled(acc)
End Sub

Private Function BuildSettingsPrompt() As String
    Dim lines As String
    Dim acc As Outlook.Account
    Dim i As Long

    lines = "Choose which mailboxes AES scans automatically." & vbCrLf & vbCrLf

    For i = 1 To Application.Session.Accounts.Count
        Set acc = Application.Session.Accounts.Item(i)
        lines = lines & i & ". "
        If IsAccountScanEnabled(acc) Then
            lines = lines & "[ON]  "
        Else
            lines = lines & "[OFF] "
        End If
        lines = lines & acc.DisplayName & vbCrLf
        lines = lines & "    " & AccountSmtpAddress(acc) & vbCrLf
    Next i

    lines = lines & vbCrLf & _
        "Enter account number to toggle ON/OFF" & vbCrLf & _
        "A = all ON   N = all OFF   S = save   C = cancel"
    BuildSettingsPrompt = lines
End Function

Private Sub LoadSettingsFromFile()
    Dim path As String
    path = GetSettingsFilePath()
    If Len(path) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        SaveSettingsToFile
        Exit Sub
    End If

    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1, False)
    Do While Not ts.AtEndOfStream
        Dim line As String
        line = Trim$(ts.ReadLine)
        If Len(line) = 0 Then GoTo NextLine
        If Left$(line, 1) = ";" Or Left$(line, 1) = "#" Then GoTo NextLine
        If InStr(line, "=") = 0 Then GoTo NextLine

        Dim storeId As String
        Dim valuePart As String
        Dim pipePos As Long
        Dim parts As Variant
        Dim enabledFlag As Boolean
        Dim responsesFlag As Boolean
        Dim inCcFlag As Boolean
        storeId = Trim$(Left$(line, InStr(line, "=") - 1))
        valuePart = Trim$(Mid$(line, InStr(line, "=") + 1))
        ' Save writes storeId|smtp=enabled,responses,in_cc; lookups use bare storeId.
        pipePos = InStr(storeId, "|")
        If pipePos > 0 Then storeId = Trim$(Left$(storeId, pipePos - 1))
        If Len(storeId) > 0 Then
            responsesFlag = True
            inCcFlag = True
            If InStr(valuePart, ",") > 0 Then
                parts = Split(valuePart, ",")
                enabledFlag = (Trim$(CStr(parts(0))) = "1" Or UCase$(Trim$(CStr(parts(0)))) = "TRUE")
                If UBound(parts) >= 1 Then
                    responsesFlag = (Trim$(CStr(parts(1))) = "1" Or UCase$(Trim$(CStr(parts(1)))) = "TRUE")
                End If
                If UBound(parts) >= 2 Then
                    inCcFlag = (Trim$(CStr(parts(2))) = "1" Or UCase$(Trim$(CStr(parts(2)))) = "TRUE")
                End If
            Else
                enabledFlag = (valuePart = "1" Or UCase$(valuePart) = "ON" Or UCase$(valuePart) = "TRUE")
            End If
            mAccountStates(storeId) = enabledFlag
            mAccountResponses(storeId) = responsesFlag
            mAccountInCc(storeId) = inCcFlag
        End If
NextLine:
    Loop
    ts.Close
End Sub

Private Sub SaveSettingsToFile()
    SyncAccountsFromOutlook

    Dim path As String
    path = GetSettingsFilePath()
    If Len(path) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.CreateTextFile(path, True)
    ts.WriteLine "; AES scan account settings"
    ts.WriteLine "; store_id|smtp=enabled,responses,in_cc  (1=on, 0=off)"
    ts.WriteLine "; Updated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Dim acc As Outlook.Account
    Dim storeId As String
    For Each acc In Application.Session.Accounts
        storeId = AccountStoreId(acc)
        If Len(storeId) > 0 Then
            Dim flag As String
            Dim respFlag As String
            Dim ccFlag As String
            If IsAccountScanEnabled(acc) Then
                flag = "1"
            Else
                flag = "0"
            End If
            If IsAccountResponsesEnabled(acc) Then
                respFlag = "1"
            Else
                respFlag = "0"
            End If
            If IsAccountInCcEnabled(acc) Then
                ccFlag = "1"
            Else
                ccFlag = "0"
            End If
            ts.WriteLine storeId & "|" & AccountSmtpAddress(acc) & "=" & flag & "," & respFlag & "," & ccFlag
        End If
    Next acc
    ts.Close
End Sub

Private Function AccountStoreId(ByVal acc As Outlook.Account) As String
    On Error Resume Next
    If Not acc.DeliveryStore Is Nothing Then
        AccountStoreId = acc.DeliveryStore.StoreID
    End If
    If Len(AccountStoreId) = 0 Then AccountStoreId = acc.StoreID
End Function

Private Function AccountSmtpAddress(ByVal acc As Outlook.Account) As String
    On Error Resume Next
    AccountSmtpAddress = acc.SmtpAddress
    If Len(AccountSmtpAddress) = 0 Then AccountSmtpAddress = acc.DisplayName
End Function
