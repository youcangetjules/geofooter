Attribute VB_Name = "MSCANDiagnostics"
'===============================================================================
' MSCANDiagnostics - Comprehensive Diagnostic and Testing Module
' Author: Julian Garrett (Rewritten)
' Purpose: Provide full diagnostic, testing, and troubleshooting capabilities
'===============================================================================
Option Explicit

Private Const MODULE_NAME As String = "MSCANDiagnostics"

'===============================================================================
' MAIN DIAGNOSTIC ENTRY POINT
'===============================================================================

' Toolbar / ribbon: interactive PySide6 dialog (per-test Run). Falls back to text report.
Public Sub ShowDiagnosticsDialog()
    On Error GoTo EH

    If TryShowPythonDiagnosticsDialog() Then
        Exit Sub
    End If

    MSCANModLogging.WriteLogWarn "ShowDiagnosticsDialog: Python UI unavailable; text report fallback."
    RunFullDiagnostics
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ShowDiagnosticsDialog error: #" & Err.Number & " - " & Err.Description
    On Error Resume Next
    RunFullDiagnostics
End Sub

Public Sub RunFullDiagnostics()
    '''Runs complete diagnostic suite and displays results (Notepad text report)
    
    On Error GoTo ErrorHandler
    
    Dim results As String
    results = ""
    
    results = results & "========================================" & vbCrLf
    results = results & "AES FULL DIAGNOSTIC REPORT" & vbCrLf
    results = results & "Generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    results = results & "========================================" & vbCrLf & vbCrLf
    
    ' System Information
    results = results & GetSystemInfo() & vbCrLf
    
    ' Outlook Environment
    results = results & GetOutlookInfo() & vbCrLf
    
    ' AES scanner / Python pipeline
    results = results & GetAesScannerInfo() & vbCrLf
    results = results & RunAesScannerTests() & vbCrLf
    
    ' AES service / watchers / auto-scan pipeline
    results = results & GetAesServiceInfo() & vbCrLf
    results = results & MSCANModQueueManager.GetAutoScanDiagnostics() & vbCrLf
    
    ' Account scan settings
    results = results & MSCANSettings.GetAccountScanDiagnostics() & vbCrLf
    
    ' Recent VBA log auto-scan lines
    results = results & GetRecentAutoScanLogExcerpt() & vbCrLf
    
    ' MSCAN Configuration
    results = results & GetConfigurationInfo() & vbCrLf
    
    ' Classification Tests
    results = results & RunClassificationTests() & vbCrLf
    
    ' Form Tests
    results = results & RunFormTests() & vbCrLf
    
    ' Event Handler Tests
    results = results & RunEventTests() & vbCrLf
    
    ' Performance Tests
    results = results & RunPerformanceTests() & vbCrLf
    
    ' Log File Status
    results = results & GetLogFileStatus() & vbCrLf
    
    results = results & "========================================" & vbCrLf
    results = results & "END OF DIAGNOSTIC REPORT" & vbCrLf
    results = results & "========================================" & vbCrLf
    
    ' Display results
    DisplayDiagnosticResults results
    
    MSCANCore.Log "Full diagnostics completed"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in RunFullDiagnostics: " & Err.Description & vbCrLf & _
           "Number: " & Err.Number, vbCritical, "MSCAN Diagnostics Error"
End Sub

Private Function TryShowPythonDiagnosticsDialog() As Boolean
    On Error GoTo EH

    Dim pythonExe As String
    Dim scriptPath As String
    Dim contextPath As String
    Dim cmd As String
    Dim shell As Object

    TryShowPythonDiagnosticsDialog = False

    pythonExe = ResolveDiagnosticsPythonExe()
    scriptPath = ResolveDiagnosticsDialogScript()
    If Len(pythonExe) = 0 Or Len(scriptPath) = 0 Then Exit Function

    contextPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_diagnostics_context.json"
    If Not WriteDiagnosticsContextJson(contextPath) Then Exit Function

    ' Launch WITHOUT waiting — a modal Wait freezes Outlook for the whole
    ' dialog session (no ItemLoad / queue / timers). Live Controls need Outlook
    ' responsive so diag commands + VBS nudges can run.
    cmd = """" & pythonExe & """ """ & scriptPath & """ --context " & _
          """" & Replace(contextPath, """", """""") & """"

    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, 1, False
    MSCANModLogging.WriteLogInfo "ShowDiagnosticsDialog: launched interactive UI (non-blocking)."
    TryShowPythonDiagnosticsDialog = True
    Exit Function

EH:
    MSCANModLogging.WriteLog "TryShowPythonDiagnosticsDialog error: #" & Err.Number & " - " & Err.Description
    TryShowPythonDiagnosticsDialog = False
End Function

Private Function WriteDiagnosticsContextJson(ByVal filePath As String) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Dim ts As Object
    Dim i As Long
    Dim tagCount As Long
    Dim first As Boolean
    Dim parent As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    parent = fso.GetParentFolderName(filePath)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then fso.CreateFolder parent
    End If

    Set ts = fso.CreateTextFile(filePath, True, False)
    ts.WriteLine "{"
    ts.WriteLine "  ""service_enabled"": " & IIf(MSCANModWatchers.IsServiceEnabled(), "true", "false") & ","
    ts.WriteLine "  ""watcher_count"": " & CLng(MSCANModWatchers.GetWatcherCount()) & ","
    ts.WriteLine "  ""queue_size"": " & CLng(MSCANModQueueManager.GetQueueSize()) & ","
    ts.WriteLine "  ""inflight_scans"": " & CLng(MSCANModule1.PendingAsyncJobCount()) & ","
    ts.WriteLine "  ""startup_quiet"": " & IIf(MSCANModQueueManager.IsStartupQuiet(), "true", "false") & ","
    ts.WriteLine "  ""log_path"": """ & JsonEscDiag(MSCANModLogging.logPath()) & ""","
    ts.Write "  ""logging"": "
    ts.Write MSCANModLogging.GetLoggingSettingsJson()
    ts.WriteLine ","
    ts.WriteLine "  ""capture_path"": """ & JsonEscDiag(MSCANModLogging.LogCapturePath()) & ""","
    ts.Write "  ""auto_scan_summary"": """
    ts.Write JsonEscDiag(MSCANModQueueManager.GetAutoScanDiagnostics())
    ts.WriteLine ""","
    ts.Write "  ""accounts_summary"": """
    ts.Write JsonEscDiag(MSCANSettings.GetAccountScanDiagnostics())
    ts.WriteLine ""","
    ts.Write "  ""recent_auto_scan_log"": """
    ts.Write JsonEscDiag(GetRecentAutoScanLogExcerpt())
    ts.WriteLine ""","
    ts.WriteLine "  ""tags"": ["

    tagCount = MSCANCore.GetTagCount()
    first = True
    For i = 1 To tagCount
        If Not first Then ts.WriteLine ","
        first = False
        ts.Write "    """ & JsonEscDiag(MSCANCore.GetTag(i)) & """"
    Next i
    ts.WriteLine ""
    ts.WriteLine "  ]"
    ts.WriteLine "}"
    ts.Close

    WriteDiagnosticsContextJson = True
    Exit Function

EH:
    WriteDiagnosticsContextJson = False
End Function

Private Function GetRecentAutoScanLogExcerpt() As String
    On Error GoTo EH
    Dim fso As Object
    Dim path As String
    Dim buf As String
    Dim hits As String
    Dim line As String
    Dim n As Long
    Dim parts() As String
    Dim i As Long
    Dim fileSize As Long
    Dim stream As Object
    Dim startAt As Long

    path = MSCANModLogging.logPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        GetRecentAutoScanLogExcerpt = "(log file missing: " & path & ")"
        Exit Function
    End If

    ' Only read the tail — full ReadAll of a multi-MB log freezes Outlook.
    fileSize = CLng(fso.GetFile(path).Size)
    startAt = fileSize - 180000
    If startAt < 0 Then startAt = 0

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    If startAt > 0 Then
        stream.Position = startAt
        ' Discard the partial first line after a mid-file seek.
        stream.ReadText -2 ' adReadLine
    End If
    buf = stream.ReadText(-1) ' adReadAll
    stream.Close

    parts = Split(buf, vbCrLf)
    hits = ""
    n = 0
    For i = UBound(parts) To LBound(parts) Step -1
        line = parts(i)
        If Len(line) = 0 Then GoTo NextLogLine
        If InStr(1, line, "NewMailEx", vbTextCompare) > 0 Or _
           InStr(1, line, "ItemAdd", vbTextCompare) > 0 Or _
           InStr(1, line, "Queued", vbTextCompare) > 0 Or _
           InStr(1, line, "ProcessQueue", vbTextCompare) > 0 Or _
           InStr(1, line, "CatchUp", vbTextCompare) > 0 Or _
           InStr(1, line, "LaunchDelayedQueueTick", vbTextCompare) > 0 Or _
           InStr(1, line, "watcher", vbTextCompare) > 0 Or _
           InStr(1, line, "account scan", vbTextCompare) > 0 Or _
           InStr(1, line, "StartAsyncGeolocationJob", vbTextCompare) > 0 Or _
           InStr(1, line, "CompleteAsyncFooter", vbTextCompare) > 0 Then
            hits = line & vbCrLf & hits
            n = n + 1
            If n >= 40 Then Exit For
        End If
NextLogLine:
    Next i

    If n = 0 Then
        GetRecentAutoScanLogExcerpt = "(no recent auto-scan lines in log tail)"
    Else
        GetRecentAutoScanLogExcerpt = "--- LAST " & n & " AUTO-SCAN LOG LINES ---" & vbCrLf & hits
    End If
    Exit Function

EH:
    GetRecentAutoScanLogExcerpt = "(error reading log: #" & Err.Number & " - " & Err.Description & ")"
End Function

Private Function JsonEscDiag(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    JsonEscDiag = t
End Function

Private Function ResolveDiagnosticsPythonExe() As String
    On Error Resume Next
    Dim fso As Object
    Dim paths As Variant
    Dim p As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    paths = Array( _
        "C:\Python313\pythonw.exe", _
        "C:\Python313\python.exe", _
        Environ$("LOCALAPPDATA") & "\Programs\Python\Python313\pythonw.exe", _
        "C:\GeoFooter\.venv\Scripts\pythonw.exe", _
        "C:\GeoFooter\.venv\Scripts\python.exe")
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            ResolveDiagnosticsPythonExe = CStr(p)
            Exit Function
        End If
    Next p
    ResolveDiagnosticsPythonExe = ""
End Function

Private Function ResolveDiagnosticsDialogScript() As String
    On Error Resume Next
    Dim fso As Object
    Dim paths As Variant
    Dim p As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    paths = Array( _
        "C:\GeoFooter\VBA\aes_diagnostics_dialog.py", _
        "C:\GeoFooter\aes_diagnostics_dialog.py", _
        Environ$("LOCALAPPDATA") & "\GeoFooter\aes_diagnostics_dialog.py")
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            ResolveDiagnosticsDialogScript = CStr(p)
            Exit Function
        End If
    Next p
    ResolveDiagnosticsDialogScript = ""
End Function

'===============================================================================
' SYSTEM INFORMATION
'===============================================================================

Private Function GetSystemInfo() As String
    Dim info As String
    info = "--- SYSTEM INFORMATION ---" & vbCrLf
    
    On Error Resume Next
    
    ' Operating System
    info = info & "OS: " & Environ$("OS") & vbCrLf
    info = info & "Computer: " & Environ$("COMPUTERNAME") & vbCrLf
    info = info & "Username: " & Environ$("USERNAME") & vbCrLf
    info = info & "User Domain: " & Environ$("USERDOMAIN") & vbCrLf
    
    ' VBA Environment
    info = info & "VBA Version: " & Application.Version & vbCrLf
    
    ' Temp path
    info = info & "Temp Path: " & Environ$("TEMP") & vbCrLf
    
    ' Current time
    info = info & "Local Time: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    
    On Error GoTo 0
    
    info = info & vbCrLf
    GetSystemInfo = info
End Function

'===============================================================================
' OUTLOOK INFORMATION
'===============================================================================

Private Function GetOutlookInfo() As String
    Dim info As String
    info = "--- OUTLOOK ENVIRONMENT ---" & vbCrLf
    
    On Error Resume Next
    
    Dim olApp As Outlook.Application
    Set olApp = Application
    
    info = info & "Outlook Version: " & olApp.Version & vbCrLf
    info = info & "Build: " & olApp.Build & vbCrLf
    info = info & "Product Code: " & olApp.ProductCode & vbCrLf
    
    ' Default profile
    Dim ns As Outlook.NameSpace
    Set ns = olApp.GetNamespace("MAPI")
    
    info = info & "Current User: " & ns.CurrentUser.Name & vbCrLf
    info = info & "Current User Address: " & ns.CurrentUser.Address & vbCrLf
    
    ' Count accounts
    Dim acctCount As Long
    acctCount = ns.Accounts.count
    info = info & "Account Count: " & acctCount & vbCrLf
    
    ' List accounts
    If acctCount > 0 Then
        Dim acct As Outlook.Account
        Dim i As Long
        For i = 1 To acctCount
            Set acct = ns.Accounts.Item(i)
            info = info & "  Account " & i & ": " & acct.displayName
            info = info & " (" & acct.SmtpAddress & ")" & vbCrLf
        Next i
    End If
    
    ' Default folders accessible?
    info = info & "Inbox Accessible: " & TestFolderAccess(olFolderInbox) & vbCrLf
    info = info & "Sent Items Accessible: " & TestFolderAccess(olFolderSentMail) & vbCrLf
    info = info & "Drafts Accessible: " & TestFolderAccess(olFolderDrafts) & vbCrLf
    info = info & "Outbox Accessible: " & TestFolderAccess(olFolderOutbox) & vbCrLf
    
    ' Explorer/Inspector count
    info = info & "Open Explorers: " & olApp.Explorers.count & vbCrLf
    info = info & "Open Inspectors: " & olApp.Inspectors.count & vbCrLf
    
    Set ns = Nothing
    Set olApp = Nothing
    
    On Error GoTo 0
    
    info = info & vbCrLf
    GetOutlookInfo = info
End Function

Private Function TestFolderAccess(folderType As OlDefaultFolders) As String
    On Error Resume Next
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    Dim fld As Outlook.MAPIFolder
    Set fld = ns.GetDefaultFolder(folderType)
    
    If Err.Number = 0 And Not fld Is Nothing Then
        TestFolderAccess = "Yes (" & fld.Items.count & " items)"
    Else
        TestFolderAccess = "No - " & Err.Description
    End If
    
    Err.Clear
    Set fld = Nothing
    Set ns = Nothing
End Function

'===============================================================================
' CONFIGURATION INFORMATION
'===============================================================================

Private Function GetConfigurationInfo() As String
    Dim info As String
    info = "--- MSCAN CONFIGURATION ---" & vbCrLf
    
    On Error Resume Next
    
    ' Classification tags
    Dim tagCount As Long
    tagCount = MSCANCore.GetTagCount()
    
    info = info & "Classification Tags Defined: " & tagCount & vbCrLf
    
    Dim i As Long
    For i = 1 To tagCount
        info = info & "  " & i & ". " & MSCANCore.GetTag(i) & vbCrLf
    Next i
    
    ' Logging status
    info = info & "Logging Enabled: " & MSCANCore.IsLoggingEnabled() & vbCrLf
    info = info & "Log File Path: " & MSCANCore.GetLogFilePath() & vbCrLf
    
    ' Check if log file exists and is writable
    Dim logPath As String
    logPath = MSCANCore.GetLogFilePath()
    
    If Len(logPath) > 0 Then
        If Dir(logPath) <> "" Then
            info = info & "Log File Exists: Yes" & vbCrLf
            info = info & "Log File Size: " & FileLen(logPath) & " bytes" & vbCrLf
        Else
            info = info & "Log File Exists: No" & vbCrLf
        End If
    End If
    
    ' Detection pattern
    info = info & "Detection Pattern: " & MSCANCore.GetDetectionPattern() & vbCrLf
    
    On Error GoTo 0
    
    info = info & vbCrLf
    GetConfigurationInfo = info
End Function

'===============================================================================
' CLASSIFICATION TESTS
'===============================================================================

Private Function RunClassificationTests() As String
    Dim info As String
    info = "--- CLASSIFICATION TESTS ---" & vbCrLf
    
    Dim passed As Long
    Dim failed As Long
    passed = 0
    failed = 0
    
    On Error Resume Next
    
    ' Get actual tags from MSCANCore
    Dim tag1 As String
    Dim tag2 As String
    tag1 = MSCANCore.GetTag(1)  ' [NR/E]
    tag2 = MSCANCore.GetTag(2)  ' [SEC1: (C) NOT RATED /EXTERNAL]
    
    ' Test 1: Detect known tag (tag1)
    If RunSingleTest("Detect tag1", _
        MSCANCore.DetectClassification(tag1 & " Test"), _
        tag1) Then
        passed = passed + 1
        info = info & "  PASS: Detect " & tag1 & " tag" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: Detect " & tag1 & " tag" & vbCrLf
    End If
    
    ' Test 2: Detect no tag
    If RunSingleTest("Detect none", _
        MSCANCore.DetectClassification("Plain subject"), _
        "") Then
        passed = passed + 1
        info = info & "  PASS: Detect no tag" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: Detect no tag" & vbCrLf
    End If
    
    ' Test 3: Has classification true
    If MSCANCore.HasClassification(tag2 & " Meeting") = True Then
        passed = passed + 1
        info = info & "  PASS: HasClassification true" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: HasClassification true" & vbCrLf
    End If
    
    ' Test 4: Has classification false
    If MSCANCore.HasClassification("Plain meeting") = False Then
        passed = passed + 1
        info = info & "  PASS: HasClassification false" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: HasClassification false" & vbCrLf
    End If
    
    ' Test 5: Apply classification BY INDEX
    Dim applied As String
    applied = MSCANCore.ApplyClassification("Test subject", 1)  ' Use index 1
    If InStr(applied, tag1) > 0 And InStr(applied, "Test subject") > 0 Then
        passed = passed + 1
        info = info & "  PASS: Apply classification by index" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: Apply classification by index" & vbCrLf
    End If
    
    ' Test 6: Remove classification
    Dim removed As String
    removed = MSCANCore.RemoveClassification(tag1 & " Test subject")
    If InStr(removed, tag1) = 0 And InStr(removed, "Test subject") > 0 Then
        passed = passed + 1
        info = info & "  PASS: Remove classification" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: Remove classification" & vbCrLf
    End If
    
    ' Test 7: Replace classification
    Dim replaced As String
    replaced = MSCANCore.ApplyClassification(tag1 & " Old subject", 2)  ' Use index 2
    If InStr(replaced, tag2) > 0 And InStr(replaced, tag1) = 0 Then
        passed = passed + 1
        info = info & "  PASS: Replace classification" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: Replace classification" & vbCrLf
    End If
    
    ' Test 8: Valid classification check
    If MSCANCore.IsValidClassification(tag1) = True Then
        passed = passed + 1
        info = info & "  PASS: IsValidClassification true" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: IsValidClassification true" & vbCrLf
    End If
    
    ' Test 9: Invalid classification check
    If MSCANCore.IsValidClassification("[INVALID]") = False Then
        passed = passed + 1
        info = info & "  PASS: IsValidClassification false" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: IsValidClassification false" & vbCrLf
    End If
    
    ' Test 10: Get classification level
    Dim level As Long
    level = MSCANCore.GetClassificationLevel(tag1)
    If level = 1 Then
        passed = passed + 1
        info = info & "  PASS: GetClassificationLevel" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: GetClassificationLevel (expected 1, got " & level & ")" & vbCrLf
    End If
    
    ' Test 11: Apply classification BY TAG STRING
    Dim appliedByTag As String
    appliedByTag = MSCANCore.ApplyClassificationByTag("Test subject", tag2)
    If InStr(appliedByTag, tag2) > 0 And InStr(appliedByTag, "Test subject") > 0 Then
        passed = passed + 1
        info = info & "  PASS: Apply classification by tag string" & vbCrLf
    Else
        failed = failed + 1
        info = info & "  FAIL: Apply classification by tag string" & vbCrLf
    End If
    
    On Error GoTo 0
    
    info = info & "Classification Tests: " & passed & " passed, " & failed & " failed" & vbCrLf
    info = info & vbCrLf
    
    RunClassificationTests = info
End Function

Private Function RunSingleTest(testName As String, actual As String, expected As String) As Boolean
    RunSingleTest = (StrComp(actual, expected, vbTextCompare) = 0)
End Function

'===============================================================================
' FORM TESTS
'===============================================================================

Private Function RunFormTests() As String
    Dim info As String
    info = "--- CLASSIFICATION DIALOG TESTS ---" & vbCrLf
    
    On Error Resume Next
    
    Dim tagCount As Long
    tagCount = MSCANCore.GetTagCount()
    If Err.Number <> 0 Then
        info = info & "  FAIL: MSCANCore.GetTagCount - " & Err.Description & vbCrLf
        Err.Clear
    ElseIf tagCount = 0 Then
        info = info & "  FAIL: No classification tags configured" & vbCrLf
    Else
        info = info & "  PASS: Classification tags available (" & tagCount & ")" & vbCrLf
        info = info & "  NOTE: Uses MSCANClassificationDialog (no UserForm import)" & vbCrLf
        info = info & "  Manual test: Alt+F8 > MSCANClassificationDialog.QuickTest" & vbCrLf
    End If
    
    On Error GoTo 0
    
    info = info & vbCrLf
    RunFormTests = info
End Function

'===============================================================================
' EVENT HANDLER TESTS
'===============================================================================

Private Function RunEventTests() As String
    Dim info As String
    info = "--- EVENT HANDLER TESTS ---" & vbCrLf
    
    On Error Resume Next
    
    info = info & "  Note: Event handlers cannot be directly tested" & vbCrLf
    info = info & "  Manual verification required for:" & vbCrLf
    info = info & "    - Application_ItemSend" & vbCrLf
    info = info & "    - Application_Startup" & vbCrLf
    info = info & "    - Application_Quit" & vbCrLf
    
    ' Test that we can create a draft email
    Dim testMail As Outlook.mailItem
    Set testMail = Application.CreateItem(olMailItem)
    
    If Err.Number <> 0 Then
        info = info & "  FAIL: Cannot create MailItem - " & Err.Description & vbCrLf
        Err.Clear
    Else
        info = info & "  PASS: MailItem creation" & vbCrLf
        
        testMail.Subject = "MSCAN Diagnostic Test"
        testMail.To = "test@example.com"
        testMail.Body = "This is a diagnostic test."
        
        info = info & "  PASS: MailItem property assignment" & vbCrLf
        
        ' Clean up without saving
        testMail.Close olDiscard
    End If
    
    Set testMail = Nothing
    
    On Error GoTo 0
    
    info = info & vbCrLf
    RunEventTests = info
End Function

'===============================================================================
' PERFORMANCE TESTS
'===============================================================================

Private Function RunPerformanceTests() As String
    Dim info As String
    info = "--- PERFORMANCE TESTS ---" & vbCrLf
    
    On Error Resume Next
    
    Dim startTime As Double
    Dim endTime As Double
    Dim iterations As Long
    iterations = 1000
    
    Dim tag1 As String
    tag1 = MSCANCore.GetTag(1)
    
    ' Test DetectClassification performance
    startTime = Timer
    Dim i As Long
    Dim dummy As String
    For i = 1 To iterations
        dummy = MSCANCore.DetectClassification(tag1 & " Test subject line here")
    Next i
    endTime = Timer
    
    info = info & "  DetectClassification (" & iterations & " iterations): "
    info = info & Format$((endTime - startTime) * 1000, "0.00") & " ms total, "
    info = info & Format$((endTime - startTime) * 1000 / iterations, "0.0000") & " ms each" & vbCrLf
    
    ' Test ApplyClassification performance - FIXED: Use index not string
    startTime = Timer
    For i = 1 To iterations
        dummy = MSCANCore.ApplyClassification("Test subject", 1)  ' Use index 1
    Next i
    endTime = Timer
    
    info = info & "  ApplyClassification (" & iterations & " iterations): "
    info = info & Format$((endTime - startTime) * 1000, "0.00") & " ms total, "
    info = info & Format$((endTime - startTime) * 1000 / iterations, "0.0000") & " ms each" & vbCrLf
    
    ' Test HasClassification performance
    startTime = Timer
    Dim dummy2 As Boolean
    For i = 1 To iterations
        dummy2 = MSCANCore.HasClassification(tag1 & " Test subject")
    Next i
    endTime = Timer
    
    info = info & "  HasClassification (" & iterations & " iterations): "
    info = info & Format$((endTime - startTime) * 1000, "0.00") & " ms total, "
    info = info & Format$((endTime - startTime) * 1000 / iterations, "0.0000") & " ms each" & vbCrLf
    
    On Error GoTo 0
    
    info = info & vbCrLf
    RunPerformanceTests = info
End Function

'===============================================================================
' LOG FILE STATUS
'===============================================================================

Private Function GetLogFileStatus() As String
    Dim info As String
    info = "--- LOG FILE STATUS ---" & vbCrLf
    
    On Error Resume Next
    
    Dim logPath As String
    logPath = MSCANCore.GetLogFilePath()
    
    If Len(logPath) = 0 Then
        info = info & "  Log path not configured" & vbCrLf
        GetLogFileStatus = info & vbCrLf
        Exit Function
    End If
    
    info = info & "  Path: " & logPath & vbCrLf
    
    If Dir(logPath) = "" Then
        info = info & "  Status: File does not exist" & vbCrLf
        
        ' Check if parent folder exists
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim parentFolder As String
        parentFolder = fso.GetParentFolderName(logPath)
        
        If fso.FolderExists(parentFolder) Then
            info = info & "  Parent folder exists: Yes" & vbCrLf
        Else
            info = info & "  Parent folder exists: No - " & parentFolder & vbCrLf
        End If
        Set fso = Nothing
    Else
        info = info & "  Status: File exists" & vbCrLf
        info = info & "  Size: " & Format$(FileLen(logPath), "#,##0") & " bytes" & vbCrLf
        info = info & "  Modified: " & Format$(FileDateTime(logPath), "yyyy-mm-dd hh:nn:ss") & vbCrLf
    End If
    
    On Error GoTo 0
    
    info = info & vbCrLf
    GetLogFileStatus = info
End Function

'===============================================================================
' DISPLAY RESULTS
'===============================================================================

Private Sub DisplayDiagnosticResults(results As String)
    On Error GoTo ErrorHandler
    
    Debug.Print results
    Debug.Print ""
    Debug.Print "================================================"
    Debug.Print "Full diagnostic results printed to Immediate Window (Ctrl+G)"
    Debug.Print "================================================"
    
    Dim reportPath As String
    reportPath = BuildDiagnosticReportPath()
    
    Dim fNum As Integer
    fNum = FreeFile
    Open reportPath For Output As #fNum
    Print #fNum, results
    Close #fNum
    
    MSCANModLogging.OpenTextFileInEditor reportPath
    
    MsgBox "AES full diagnostics complete." & vbCrLf & vbCrLf & _
           SummarizeDiagnosticResults(results) & vbCrLf & vbCrLf & _
           "Full report opened:" & vbCrLf & reportPath, _
           vbInformation, "AES Diagnostics"
    Exit Sub
    
ErrorHandler:
    MsgBox "Diagnostic report could not be saved. Results are in the Immediate Window (Ctrl+G)." & vbCrLf & _
           Err.Description, vbExclamation, "AES Diagnostics"
End Sub

Private Function BuildDiagnosticReportPath() As String
    Dim baseDir As String
    baseDir = Environ$("LOCALAPPDATA") & "\GeoFooter\Diagnostics"
    
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir
    If Not fso.FolderExists(baseDir) Then baseDir = Environ$("TEMP")
    
    BuildDiagnosticReportPath = baseDir & "\AES_Diagnostic_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
End Function

Private Function SummarizeDiagnosticResults(ByVal results As String) As String
    Dim passCount As Long
    Dim failCount As Long
    Dim lines As Variant
    Dim i As Long
    Dim line As String
    
    lines = Split(results, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        line = CStr(lines(i))
        If InStr(1, line, "PASS:", vbTextCompare) > 0 Then passCount = passCount + 1
        If InStr(1, line, "FAIL:", vbTextCompare) > 0 Then failCount = failCount + 1
    Next i
    
    SummarizeDiagnosticResults = "Tests passed: " & passCount & "   Failed: " & failCount
End Function

Private Function GetAesScannerInfo() As String
    Dim info As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    info = "--- AES SCANNER / PYTHON ---" & vbCrLf
    info = info & "Python (default): C:\Python313\python.exe — " & FileExistsLabel(fso, "C:\Python313\python.exe") & vbCrLf
    info = info & "geolocate_headers.py (VBA): C:\GeoFooter\VBA\geolocate_headers.py — " & FileExistsLabel(fso, "C:\GeoFooter\VBA\geolocate_headers.py") & vbCrLf
    info = info & "geolocate_headers.py (root): C:\GeoFooter\geolocate_headers.py — " & FileExistsLabel(fso, "C:\GeoFooter\geolocate_headers.py") & vbCrLf
    info = info & "guri.py: C:\GeoFooter\guri.py — " & FileExistsLabel(fso, "C:\GeoFooter\guri.py") & vbCrLf
    info = info & "guri_postgres_config.json: C:\GeoFooter\guri_postgres_config.json — " & FileExistsLabel(fso, "C:\GeoFooter\guri_postgres_config.json") & vbCrLf
    info = info & "guri_mysql_config.json (legacy): C:\GeoFooter\guri_mysql_config.json — " & FileExistsLabel(fso, "C:\GeoFooter\guri_mysql_config.json") & vbCrLf
    info = info & "Output folder: C:\GeoFooter\output — " & FolderExistsLabel(fso, "C:\GeoFooter\output") & vbCrLf
    info = info & "VBA log: " & MSCANModLogging.logPath() & " — " & FileExistsLabel(fso, MSCANModLogging.logPath()) & vbCrLf
    info = info & vbCrLf
    GetAesScannerInfo = info
End Function

Private Function RunAesScannerTests() As String
    Dim info As String
    Dim passed As Long
    Dim failed As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    info = "--- AES SCANNER TESTS ---" & vbCrLf
    
    If fso.FileExists("C:\Python313\python.exe") Then
        info = info & "  PASS: Python executable found" & vbCrLf
        passed = passed + 1
    Else
        info = info & "  FAIL: Python executable not found at C:\Python313\python.exe" & vbCrLf
        failed = failed + 1
    End If
    
    If fso.FileExists("C:\GeoFooter\VBA\geolocate_headers.py") Or fso.FileExists("C:\GeoFooter\geolocate_headers.py") Then
        info = info & "  PASS: geolocate_headers.py found" & vbCrLf
        passed = passed + 1
    Else
        info = info & "  FAIL: geolocate_headers.py not found" & vbCrLf
        failed = failed + 1
    End If
    
    If fso.FileExists("C:\GeoFooter\guri.py") Then
        info = info & "  PASS: guri.py found" & vbCrLf
        passed = passed + 1
    Else
        info = info & "  FAIL: guri.py not found" & vbCrLf
        failed = failed + 1
    End If
    
    If MSCANModWatchers.IsServiceEnabled() Then
        info = info & "  PASS: AES automatic scanning is ON" & vbCrLf
        passed = passed + 1
    Else
        info = info & "  FAIL: AES automatic scanning is OFF" & vbCrLf
        failed = failed + 1
    End If
    
    info = info & "AES scanner tests: " & passed & " passed, " & failed & " failed" & vbCrLf & vbCrLf
    RunAesScannerTests = info
End Function

Private Function GetAesServiceInfo() As String
    Dim info As String
    info = "--- AES SERVICE / WATCHERS ---" & vbCrLf
    info = info & "Service enabled: " & MSCANModWatchers.IsServiceEnabled() & vbCrLf
    info = info & "Watchers active: " & MSCANModWatchers.AreWatchersActive() & vbCrLf
    info = info & "Watcher count: " & MSCANModWatchers.GetWatcherCount() & vbCrLf
    info = info & "Queue size: " & MSCANModQueueManager.GetQueueSize() & vbCrLf
    info = info & "In-flight scans: " & MSCANModule1.PendingAsyncJobCount() & vbCrLf
    info = info & "Startup quiet: " & MSCANModQueueManager.IsStartupQuiet() & vbCrLf
    info = info & vbCrLf
    GetAesServiceInfo = info
End Function

Private Function FileExistsLabel(ByVal fso As Object, ByVal path As String) As String
    On Error Resume Next
    If Len(path) = 0 Then
        FileExistsLabel = "missing"
    ElseIf fso.FileExists(path) Then
        FileExistsLabel = "found"
    Else
        FileExistsLabel = "missing"
    End If
End Function

Private Function FolderExistsLabel(ByVal fso As Object, ByVal path As String) As String
    On Error Resume Next
    If fso.FolderExists(path) Then
        FolderExistsLabel = "found"
    Else
        FolderExistsLabel = "missing"
    End If
End Function

'===============================================================================
' QUICK TESTS
'===============================================================================

Public Sub QuickTestClassification()
    '''Quick test of classification detection
    Dim testSubject As String
    testSubject = InputBox("Enter a subject line to test:", "MSCAN Quick Test", _
                           MSCANCore.GetTag(1) & " Test message")
    
    If Len(testSubject) = 0 Then Exit Sub
    
    Dim detected As String
    detected = MSCANCore.DetectClassification(testSubject)
    
    Dim hasClass As Boolean
    hasClass = MSCANCore.HasClassification(testSubject)
    
    MsgBox "Subject: " & testSubject & vbCrLf & vbCrLf & _
           "Has Classification: " & hasClass & vbCrLf & _
           "Detected Tag: [" & detected & "]", _
           vbInformation, "MSCAN Quick Test"
End Sub

Public Sub QuickTestApply()
    '''Quick test of applying classification
    Dim testSubject As String
    testSubject = InputBox("Enter a subject line:", "MSCAN Quick Test", "Meeting tomorrow")
    
    If Len(testSubject) = 0 Then Exit Sub
    
    Dim tagCount As Long
    tagCount = MSCANCore.GetTagCount()
    
    Dim tagList As String
    Dim i As Long
    For i = 1 To tagCount
        tagList = tagList & i & ". " & MSCANCore.GetTag(i) & vbCrLf
    Next i
    
    Dim selection As String
    selection = InputBox("Select a classification:" & vbCrLf & vbCrLf & tagList, _
                         "MSCAN Quick Test", "1")
    
    If Len(selection) = 0 Then Exit Sub
    
    Dim selIndex As Long
    selIndex = Val(selection)
    
    If selIndex < 1 Or selIndex > tagCount Then
        MsgBox "Invalid selection", vbExclamation
        Exit Sub
    End If
    
    Dim newSubject As String
    newSubject = MSCANCore.ApplyClassification(testSubject, selIndex)
    
    MsgBox "Original: " & testSubject & vbCrLf & vbCrLf & _
           "Classified: " & newSubject, _
           vbInformation, "MSCAN Quick Test"
End Sub

Public Sub QuickTestForm()
    '''Quick test of the classification dialog (no UserForm required)
    MSCANClassificationDialog.QuickTest
End Sub

Public Sub ViewLogFile()
    '''Opens the log file in Notepad
    On Error Resume Next
    
    Dim logPath As String
    logPath = MSCANCore.GetLogFilePath()
    
    If Len(logPath) = 0 Then
        MsgBox "Log file path not configured", vbExclamation
        Exit Sub
    End If
    
    If Dir(logPath) = "" Then
        MsgBox "Log file does not exist: " & logPath, vbExclamation
        Exit Sub
    End If
    
    MSCANModLogging.OpenTextFileInEditor logPath
End Sub

Public Sub ClearLogFile()
    '''Clears the log file after confirmation
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to clear the log file?", _
                      vbQuestion + vbYesNo, "MSCAN")
    
    If response = vbNo Then Exit Sub
    
    On Error Resume Next
    
    Dim logPath As String
    logPath = MSCANCore.GetLogFilePath()
    
    If Len(logPath) = 0 Then
        MsgBox "Log file path not configured", vbExclamation
        Exit Sub
    End If
    
    Dim fNum As Integer
    fNum = FreeFile
    Open logPath For Output As #fNum
    Print #fNum, "=== LOG CLEARED " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    Close #fNum
    
    If Err.Number = 0 Then
        MsgBox "Log file cleared", vbInformation
    Else
        MsgBox "Error clearing log: " & Err.Description, vbExclamation
    End If
End Sub

Public Sub ValidateConfiguration()
    '''Validates the current MSCAN configuration
    Dim issues As String
    issues = ""
    
    Dim warnings As String
    warnings = ""
    
    On Error Resume Next
    
    ' Check classification tags
    Dim tagCount As Long
    tagCount = MSCANCore.GetTagCount()
    
    If Err.Number <> 0 Then
        issues = issues & "- Cannot retrieve classification tag count: " & Err.Description & vbCrLf
        Err.Clear
    ElseIf tagCount = 0 Then
        issues = issues & "- No classification tags defined" & vbCrLf
    Else
        Dim i As Long
        For i = 1 To tagCount
            Dim tagStr As String
            tagStr = MSCANCore.GetTag(i)
            If Left$(tagStr, 1) <> "[" Or Right$(tagStr, 1) <> "]" Then
                warnings = warnings & "- Tag may be malformed: " & tagStr & vbCrLf
            End If
        Next i
    End If
    
    ' Check log file path
    Dim logPath As String
    logPath = MSCANCore.GetLogFilePath()
    
    If Len(logPath) > 0 Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim parentFolder As String
        parentFolder = fso.GetParentFolderName(logPath)
        
        If Not fso.FolderExists(parentFolder) Then
            warnings = warnings & "- Log folder does not exist: " & parentFolder & vbCrLf
        End If
        
        Set fso = Nothing
    End If
    
    On Error GoTo 0
    
    ' Display results
    Dim msg As String
    
    If Len(issues) = 0 And Len(warnings) = 0 Then
        msg = "Configuration is valid. No issues found."
        MsgBox msg, vbInformation, "MSCAN Configuration"
    Else
        msg = ""
        If Len(issues) > 0 Then
            msg = "ISSUES (must fix):" & vbCrLf & issues & vbCrLf
        End If
        If Len(warnings) > 0 Then
            msg = msg & "WARNINGS:" & vbCrLf & warnings
        End If
        MsgBox msg, IIf(Len(issues) > 0, vbExclamation, vbInformation), "MSCAN Configuration"
    End If
End Sub

Public Sub QuickDiagnosticTest()
    '''Simple test that definitely shows output
    
    Dim msg As String
    
    msg = "=== MSCAN QUICK DIAGNOSTIC ===" & vbCrLf & vbCrLf
    
    ' Test 1: Can we get tags?
    On Error Resume Next
    Dim tagCount As Long
    tagCount = MSCANCore.GetTagCount()
    
    If Err.Number = 0 Then
        msg = msg & "? Tag Count: " & tagCount & vbCrLf
        
        Dim i As Long
        For i = 1 To tagCount
            msg = msg & "  Tag " & i & ": " & MSCANCore.GetTag(i) & vbCrLf
        Next i
    Else
        msg = msg & "? ERROR getting tags: " & Err.Description & vbCrLf
    End If
    Err.Clear
    
    msg = msg & vbCrLf
    
    ' Test 2: Logging enabled?
    msg = msg & "Logging Enabled: " & MSCANCore.IsLoggingEnabled() & vbCrLf
    msg = msg & "Log Path: " & MSCANCore.GetLogFilePath() & vbCrLf
    
    ' Test 3: Does log file exist?
    Dim logPath As String
    logPath = MSCANCore.GetLogFilePath()
    If Dir(logPath) <> "" Then
        msg = msg & "Log File Exists: Yes" & vbCrLf
        msg = msg & "Log Size: " & FileLen(logPath) & " bytes" & vbCrLf
    Else
        msg = msg & "Log File Exists: No" & vbCrLf
    End If
    
    msg = msg & vbCrLf
    
    ' Test 4: Can we detect a classification?
    Dim testSubject As String
    testSubject = MSCANCore.GetTag(1) & " Test Email"
    Dim detected As String
    detected = MSCANCore.DetectClassification(testSubject)
    
    msg = msg & "Test Subject: " & testSubject & vbCrLf
    msg = msg & "Detected: " & detected & vbCrLf
    msg = msg & "Match: " & (detected = MSCANCore.GetTag(1)) & vbCrLf
    
    ' Always show in message box
    MsgBox msg, vbInformation, "MSCAN Quick Diagnostic"
    
    ' Also print to debug
    Debug.Print msg
End Sub

