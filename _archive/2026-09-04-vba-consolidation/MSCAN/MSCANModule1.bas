Attribute VB_Name = "MSCANModule1"
Option Explicit

' Module-level state
Private m_LastAttachmentScanDir As String
Private m_LastDeepPayloadDir As String
' Async geo jobs: jobId -> Dictionary(EntryID, FooterPath, Mode, AttachDir, PayloadDir, Subject)
Private m_AsyncJobs As Object
Private m_AsyncJobSeq As Long
Private m_ReconcilingJobs As Boolean
Private m_InNudge As Boolean
Private m_NudgeWiredLogged As Boolean

'===============================================================================
' CONFIGURATION
' - These constants serve as default fallbacks.
'===============================================================================
Private Const PYTHON_EXE_DEFAULT As String = "C:\Python313\python.exe"
Private Const PYTHON_SCRIPT_DEFAULT As String = "C:\GeoFooter\VBA\geolocate_headers.py"
Private Const AES_SCANNED_PROP As String = "AESScanned"
' Status strip: Outlook refuses to load file:// images in a message, so the
' rendered PNG is attached as a hidden inline part and referenced by cid.
Private Const AES_BANNER_CID As String = "aesstatusbanner"
Private Const AES_BANNER_MARKER As String = "<!-- AES-Banner-Img: "
Private Const AES_BANNER_FILE_PREFIX As String = "aes_status_"

'===============================================================================
' Busy UI: pending async geo/deep jobs (yellow AES control while processing).
'===============================================================================
Public Function PendingAsyncJobCount() As Long
    On Error Resume Next
    EnsureAsyncJobs
    ReconcileAsyncJobs
    If m_AsyncJobs Is Nothing Then
        PendingAsyncJobCount = 0
    Else
        PendingAsyncJobCount = CLng(m_AsyncJobs.Count)
    End If
End Function

' Safety net: the external VBS callback can fail (e.g. Outlook restarted, bridge
' methods not yet exposed). Whenever the busy UI is refreshed, sweep pending jobs:
' if the Python output file already exists, finish the job locally; if the job is
' older than the timeout with no output, fail it so yellow never sticks forever.
Public Sub ReconcileAsyncJobs()
    Const JOB_TIMEOUT_SEC As Long = 300
    On Error GoTo EH
    If m_ReconcilingJobs Then Exit Sub
    If m_AsyncJobs Is Nothing Then Exit Sub
    If m_AsyncJobs.Count = 0 Then Exit Sub
    m_ReconcilingJobs = True

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim keys As Variant
    keys = m_AsyncJobs.keys

    Dim i As Long
    Dim job As Object
    For i = LBound(keys) To UBound(keys)
        If m_AsyncJobs.Exists(keys(i)) Then
            Set job = m_AsyncJobs(keys(i))
            If fso.FileExists(CStr(job("FooterPath"))) Then
                ' Wait until the file has content — Python creates/truncates on open
                ' before writing, so FileExists alone can race an empty file.
                If fso.GetFile(CStr(job("FooterPath"))).Size > 0 Then
                    If JobOutputLooksReady(job) Then
                        MSCANModLogging.WriteLog "ReconcileAsyncJobs: output found, completing job " & keys(i)
                        CompleteAsyncFooter CStr(keys(i))
                    End If
                End If
            ElseIf fso.FileExists(CStr(job("FooterPath")) & ".fail") Then
                MSCANModLogging.WriteLog "ReconcileAsyncJobs: fail marker found for job " & keys(i)
                On Error Resume Next
                fso.DeleteFile CStr(job("FooterPath")) & ".fail", True
                On Error GoTo EH
                FailAsyncFooter CStr(keys(i))
            ElseIf job.Exists("StartedAt") Then
                If DateDiff("s", CDate(job("StartedAt")), Now) > JOB_TIMEOUT_SEC Then
                    MSCANModLogging.WriteLog "ReconcileAsyncJobs: job " & keys(i) & " timed out, failing."
                    FailAsyncFooter CStr(keys(i))
                End If
            End If
        End If
    Next i

    m_ReconcilingJobs = False
    Exit Sub
EH:
    m_ReconcilingJobs = False
    MSCANModLogging.WriteLog "ReconcileAsyncJobs error: #" & Err.Number & " - " & Err.Description
End Sub

Public Function IsScanBusy() As Boolean
    IsScanBusy = (PendingAsyncJobCount() > 0)
End Function

' Deep jobs must contain report_url= (or a path); compact/full need HTML-ish content.
Private Function JobOutputLooksReady(ByVal job As Object) As Boolean
    On Error GoTo EH
    JobOutputLooksReady = False
    Dim path As String
    Dim mode As String
    Dim raw As String
    path = CStr(job("FooterPath"))
    mode = ""
    If job.Exists("Mode") Then mode = LCase$(CStr(job("Mode")))
    raw = ReadTextUtf8(path)
    If Len(raw) = 0 Then raw = ReadTextFileUtf8(path)
    If Len(Trim$(raw)) = 0 Then Exit Function
    If mode = "deep" Then
        JobOutputLooksReady = (InStr(1, raw, "report_url=", vbTextCompare) > 0 Or _
                               InStr(1, raw, "file:", vbTextCompare) > 0 Or _
                               InStr(1, raw, "deepscan_", vbTextCompare) > 0)
    Else
        JobOutputLooksReady = (InStr(1, raw, "<", vbBinaryCompare) > 0 Or Len(raw) > 32)
    End If
    Exit Function
EH:
    JobOutputLooksReady = False
End Function

' True for normal mail AND Outlook report items (read receipts / ReadNotify
' IPNRN notifications: MessageClass Report.IPM.Note.IPNRN, Class=olReport).
Public Function IsScannableItem(ByVal item As Object) As Boolean
    On Error Resume Next
    IsScannableItem = False
    If item Is Nothing Then Exit Function
    If TypeOf item Is Outlook.MailItem Then
        IsScannableItem = True
        Exit Function
    End If
    If TypeOf item Is Outlook.ReportItem Then
        IsScannableItem = True
        Exit Function
    End If
    Dim mc As String
    mc = UCase$(CStr(item.MessageClass))
    If Len(mc) > 0 Then
        If Left$(mc, 11) = "REPORT.IPM." Then IsScannableItem = True
    End If
End Function

' Entry point for Application_ItemLoad (ThisOutlookSession). Helper VBS scripts
' cannot call VBA directly in this Outlook build, so they read one inbox item,
' which fires ItemLoad in-process and lands here. Every branch is a cheap no-op
' when there is no pending work, so frequent firing is fine.
Public Sub NudgeAsyncWork()
    On Error Resume Next
    If m_InNudge Then Exit Sub
    m_InNudge = True
    If Not m_NudgeWiredLogged Then
        m_NudgeWiredLogged = True
        MSCANModLogging.WriteLog "NudgeAsyncWork: ItemLoad wiring active."
    End If
    ReconcileAsyncJobs
    MSCANModQueueManager.NudgeQueueWork
    MSCANModStatus.RestorePersistentStatusIfDue
    m_InNudge = False
End Sub

Private Sub NotifyScanBusyUi()
    On Error Resume Next
    MSCANAppBootstrap.RefreshScanBusyUI
End Sub

'===============================================================================
' UTILITY: Returns a writable base directory for temporary files.
'===============================================================================
Private Function GetBaseDir() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Prioritize user-specific, non-admin directories over system-wide ones.
    Dim paths As Variant: paths = Array( _
        Environ$("LOCALAPPDATA") & "\GeoFooter", _
        Environ$("TEMP") & "\GeoFooter", _
        "C:\GeoFooter") ' C:\ is a last resort, often restricted.
        
    Dim p As Variant
    For Each p In paths
        If EnsureFolderExists(p) Then
            GetBaseDir = p
            Exit Function
        End If
    Next p
    
    MSCANModLogging.WriteLog "GetBaseDir: CRITICAL - Could not create or access any base directory."
    MsgBox "Critical Error: The add-in could not find or create a writable working directory." & vbCrLf & _
           "Please check permissions for %LOCALAPPDATA% or %TEMP%.", vbCritical, "AES (Aliniant Email Scanner)"
    GetBaseDir = ""
End Function

'===============================================================================
' UTILITY: Dynamic Python path resolution.
'===============================================================================
Private Function GetPythonExe() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim paths As Variant: paths = Array( _
        PYTHON_EXE_DEFAULT, _
        Environ$("USERPROFILE") & "\AppData\Local\Programs\Python\Python313\python.exe", _
        Environ$("ProgramFiles") & "\Python\python.exe")
        
    Dim p As Variant
    For Each p In paths
        If fso.FileExists(p) Then
            GetPythonExe = p
            Exit Function
        End If
    Next p
    
    MSCANModLogging.WriteLog "GetPythonExe: Python executable (python.exe) not found in any checked locations."
    GetPythonExe = ""
End Function

Private Function GetPythonScript() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim paths As Variant: paths = Array( _
        PYTHON_SCRIPT_DEFAULT, _
        "C:\GeoFooter\geolocate_headers.py", _
        GetBaseDir() & "\geolocate_headers.py")

    Dim p As Variant
    For Each p In paths
        If fso.FileExists(p) Then
            GetPythonScript = p
            Exit Function
        End If
    Next p
    
    MSCANModLogging.WriteLog "GetPythonScript: Python script (geolocate_headers.py) not found."
    GetPythonScript = ""
End Function

'===============================================================================
' MAIN: The primary entry point for processing an email.
'===============================================================================

Public Function GetLastAttachmentScanDir() As String
    GetLastAttachmentScanDir = m_LastAttachmentScanDir
End Function

Public Sub ProcessEmail(ByVal mail As Object)
    On Error GoTo EH

    If Not WaitForMailReady(mail) Then
        MSCANModLogging.WriteLog "ProcessEmail: Mail not ready, will retry later: " & SafeSubject(mail)
        MSCANModQueueManager.QueueMailForProcessing mail
        Exit Sub
    End If

    If MailAlreadyHasFooter(mail) Then
        MarkAesScanned mail
        MSCANModLogging.WriteLog "ProcessEmail: Footer already present, skipping: " & SafeSubject(mail)
        Exit Sub
    End If

    MSCANModLogging.WriteLog "ProcessEmail: Starting async compact scan for subject: " & SafeSubject(mail)

    Dim savedFilePath As String
    savedFilePath = ExportHeadersForGeolocation(mail)

    If savedFilePath = "" Then
        MSCANModLogging.WriteLog "ProcessEmail: Could not export headers for geolocation: " & SafeSubject(mail)
        Exit Sub
    End If

    ' Python runs in a separate WScript process so Outlook UI (drafts) stay responsive.
    If Not StartAsyncGeolocationJob(mail, savedFilePath, "compact") Then
        MSCANModLogging.WriteLog "ProcessEmail: Failed to start async geolocation for: " & SafeSubject(mail)
        CleanupAttachmentScanDir m_LastAttachmentScanDir
        m_LastAttachmentScanDir = ""
    End If
    Exit Sub
EH:
    MSCANModLogging.WriteLog "ProcessEmail FATAL error: #" & Err.Number & " - " & Err.Description & " for subject: " & SafeSubject(mail)
End Sub

' Append or replace with the full AES analysis footer (Complete Footer button).
Public Sub ProcessCompleteFooter(ByVal mail As Object)
    On Error GoTo EH

    If Not WaitForMailReady(mail) Then
        MSCANModLogging.WriteLog "ProcessCompleteFooter: Mail not ready: " & SafeSubject(mail)
        Exit Sub
    End If

    If MailHasFullFooter(mail) Then
        MSCANModLogging.WriteLog "ProcessCompleteFooter: Full footer already present - " & SafeSubject(mail)
        Exit Sub
    End If

    MSCANModLogging.WriteLog "ProcessCompleteFooter: Starting async full footer for: " & SafeSubject(mail)

    Dim savedFilePath As String
    savedFilePath = ExportHeadersForGeolocation(mail)

    If savedFilePath = "" Then
        MSCANModLogging.WriteLog "ProcessCompleteFooter: Could not export headers for: " & SafeSubject(mail)
        Exit Sub
    End If

    If Not StartAsyncGeolocationJob(mail, savedFilePath, "full") Then
        MSCANModLogging.WriteLog "ProcessCompleteFooter: Failed to start async geolocation for: " & SafeSubject(mail)
        CleanupAttachmentScanDir m_LastAttachmentScanDir
        m_LastAttachmentScanDir = ""
    End If
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ProcessCompleteFooter FATAL error: #" & Err.Number & " - " & Err.Description & " for subject: " & SafeSubject(mail)
    CleanupAttachmentScanDir m_LastAttachmentScanDir
    m_LastAttachmentScanDir = ""
End Sub

' Ribbon-only Deep Scan: never mutates the email footer; opens an HTML report when done.
Public Sub ProcessDeepScan(ByVal mail As Object)
    On Error GoTo EH

    If Not WaitForMailReady(mail) Then
        MSCANModLogging.WriteLog "ProcessDeepScan: Mail not ready: " & SafeSubject(mail)
        Exit Sub
    End If

    MSCANModLogging.WriteLog "ProcessDeepScan: Starting async deep scan for: " & SafeSubject(mail)

    Dim savedFilePath As String
    savedFilePath = ExportHeadersForGeolocation(mail, True)

    If savedFilePath = "" Then
        MSCANModLogging.WriteLog "ProcessDeepScan: Could not export headers for: " & SafeSubject(mail)
        CleanupAttachmentScanDir m_LastAttachmentScanDir
        CleanupAttachmentScanDir m_LastDeepPayloadDir
        m_LastAttachmentScanDir = ""
        m_LastDeepPayloadDir = ""
        Exit Sub
    End If

    If Not StartAsyncGeolocationJob(mail, savedFilePath, "deep") Then
        MSCANModLogging.WriteLog "ProcessDeepScan: Failed to start async deep scan for: " & SafeSubject(mail)
        CleanupAttachmentScanDir m_LastAttachmentScanDir
        CleanupAttachmentScanDir m_LastDeepPayloadDir
        m_LastAttachmentScanDir = ""
        m_LastDeepPayloadDir = ""
    End If
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ProcessDeepScan FATAL error: #" & Err.Number & " - " & Err.Description & " for subject: " & SafeSubject(mail)
    CleanupAttachmentScanDir m_LastAttachmentScanDir
    CleanupAttachmentScanDir m_LastDeepPayloadDir
    m_LastAttachmentScanDir = ""
    m_LastDeepPayloadDir = ""
End Sub

' Called from aes_geo_job_*.vbs after Python finishes (outside Outlook's UI wait).
Public Sub CompleteAsyncFooter(ByVal jobId As String)
    On Error GoTo EH

    EnsureAsyncJobs
    If Len(jobId) = 0 Or Not m_AsyncJobs.Exists(jobId) Then
        MSCANModLogging.WriteLog "CompleteAsyncFooter: unknown jobId=" & jobId
        Exit Sub
    End If

    Dim job As Object
    Set job = m_AsyncJobs(jobId)
    m_AsyncJobs.Remove jobId

    Dim entryId As String
    Dim footerPath As String
    Dim mode As String
    Dim attachDir As String
    Dim payloadDir As String
    Dim subjectHint As String
    entryId = CStr(job("EntryID"))
    footerPath = CStr(job("FooterPath"))
    mode = CStr(job("Mode"))
    attachDir = CStr(job("AttachDir"))
    If job.Exists("PayloadDir") Then payloadDir = CStr(job("PayloadDir")) Else payloadDir = ""
    subjectHint = CStr(job("Subject"))

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(footerPath) Then
        MSCANModLogging.WriteLog "CompleteAsyncFooter: output missing for job " & jobId & " subject=" & subjectHint
        CleanupAttachmentScanDir attachDir
        CleanupAttachmentScanDir payloadDir
        MarkJobApplied footerPath
        NotifyScanBusyUi
        Exit Sub
    End If

    ' Deep Scan: report only — never modify the email body/footer.
    ' Always popup when done so the user can open/display the HTML report.
    If LCase$(mode) = "deep" Then
        Dim reportTarget As String
        Dim deepPrompt As String
        reportTarget = ReadDeepScanReportTarget(footerPath)
        If Len(reportTarget) = 0 Then
            reportTarget = FindLatestDeepScanReport()
            If Len(reportTarget) > 0 Then
                MSCANModLogging.WriteLog "CompleteAsyncFooter: deep job file missing report_url; falling back to latest report " & reportTarget
            End If
        End If
        If Len(reportTarget) > 0 Then
            deepPrompt = "AES Deep Scan is complete." & vbCrLf & vbCrLf & _
                         "Subject: " & Left$(subjectHint, 100) & vbCrLf & vbCrLf & _
                         "Open the report now?"
            If MsgBox(deepPrompt, vbInformation + vbYesNo + vbDefaultButton1, "AES Deep Scan") = vbYes Then
                OpenDeepScanReport reportTarget
                MSCANModStatus.ShowStatus "AES Deep Scan report opened: " & Left$(subjectHint, 50)
                MSCANModLogging.WriteLog "CompleteAsyncFooter: deep report opened for job " & jobId & " -> " & reportTarget
            Else
                MSCANModStatus.ShowStatus "AES Deep Scan ready (report not opened): " & Left$(subjectHint, 40)
                MSCANModLogging.WriteLog "CompleteAsyncFooter: deep report ready; user declined open for job " & jobId & " -> " & reportTarget
            End If
        Else
            MsgBox "AES Deep Scan finished, but the report could not be found." & vbCrLf & vbCrLf & _
                   "Subject: " & Left$(subjectHint, 100), vbExclamation, "AES Deep Scan"
            MSCANModStatus.ShowStatus "AES Deep Scan finished but report path missing: " & Left$(subjectHint, 40)
            MSCANModLogging.WriteLog "CompleteAsyncFooter: deep mode but no report URL in " & footerPath
        End If
        CleanupAttachmentScanDir attachDir
        CleanupAttachmentScanDir payloadDir
        MarkJobApplied footerPath
        NotifyScanBusyUi
        Exit Sub
    End If

    Dim mail As Object
    Set mail = ResolveMailByEntryID(entryId)
    If mail Is Nothing Then
        MSCANModLogging.WriteLog "CompleteAsyncFooter: mail not found EntryID for job " & jobId & " subject=" & subjectHint
        CleanupAttachmentScanDir attachDir
        CleanupAttachmentScanDir payloadDir
        MarkJobApplied footerPath
        NotifyScanBusyUi
        Exit Sub
    End If

    MSCANModLogging.WriteLog "CompleteAsyncFooter: applying " & mode & " footer to: " & SafeSubject(mail)

    Dim applied As Boolean
    applied = ApplyFooterToMail(mail, footerPath, mode)

    If Not applied Then
        ' Typical cause: save conflict (#-2147221239 "message has been
        ' changed") because the reading pane / a nudge touched the item.
        ' Re-resolve a fresh object and retry once — never keep saving a
        ' stale object (that is how mails get flattened to plain text).
        MSCANModLogging.WriteLog "CompleteAsyncFooter: apply failed; retrying with fresh item for job " & jobId
        Set mail = ResolveMailByEntryID(entryId)
        If Not mail Is Nothing Then
            applied = ApplyFooterToMail(mail, footerPath, mode)
        End If
    End If

    If applied Then
        StripSensitiveHeadersFromMail mail
    Else
        MSCANModLogging.WriteLog "CompleteAsyncFooter: footer NOT applied after retry for job " & jobId & " subject=" & subjectHint
    End If
    CleanupAttachmentScanDir attachDir
    CleanupAttachmentScanDir payloadDir
    MSCANModStatus.ShowStatus "AES footer ready: " & Left$(SafeSubject(mail), 50)
    MarkJobApplied footerPath
    NotifyScanBusyUi
    Exit Sub

EH:
    MSCANModLogging.WriteLog "CompleteAsyncFooter error: #" & Err.Number & " - " & Err.Description & " jobId=" & jobId
    On Error Resume Next
    If Len(footerPath) > 0 Then MarkJobApplied footerPath
    NotifyScanBusyUi
End Sub

Public Sub FailAsyncFooter(ByVal jobId As String)
    On Error Resume Next
    EnsureAsyncJobs
    If Len(jobId) = 0 Or Not m_AsyncJobs.Exists(jobId) Then Exit Sub
    Dim job As Object
    Set job = m_AsyncJobs(jobId)
    m_AsyncJobs.Remove jobId
    Dim mode As String
    Dim footerPath As String
    mode = ""
    footerPath = ""
    If job.Exists("Mode") Then mode = CStr(job("Mode"))
    If job.Exists("FooterPath") Then footerPath = CStr(job("FooterPath"))
    MSCANModLogging.WriteLog "FailAsyncFooter: Python failed for job " & jobId & " mode=" & mode & " subject=" & CStr(job("Subject"))
    CleanupAttachmentScanDir CStr(job("AttachDir"))
    If job.Exists("PayloadDir") Then CleanupAttachmentScanDir CStr(job("PayloadDir"))
    If LCase$(mode) = "deep" Then
        MsgBox "AES Deep Scan failed." & vbCrLf & vbCrLf & _
               "Subject: " & Left$(CStr(job("Subject")), 100), vbExclamation, "AES Deep Scan"
        MSCANModStatus.ShowStatus "AES Deep Scan failed: " & Left$(CStr(job("Subject")), 50)
    Else
        MSCANModStatus.ShowStatus "AES scan failed: " & Left$(CStr(job("Subject")), 50)
    End If
    If Len(footerPath) > 0 Then MarkJobApplied footerPath
    NotifyScanBusyUi
End Sub

' Sidecar for aes_geo_job_*.vbs: stop the post-Python nudge loop once VBA finishes.
Private Sub MarkJobApplied(ByVal footerPath As String)
    On Error Resume Next
    If Len(footerPath) = 0 Then Exit Sub
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(footerPath & ".applied", True)
    ts.WriteLine "applied=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ts.Close
End Sub

Private Sub ClearJobSidecars(ByVal footerPath As String)
    On Error Resume Next
    If Len(footerPath) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(footerPath & ".applied") Then fso.DeleteFile footerPath & ".applied", True
    If fso.FileExists(footerPath & ".fail") Then fso.DeleteFile footerPath & ".fail", True
End Sub

Private Function ReadDeepScanReportTarget(ByVal outputPath As String) As String
    On Error GoTo EH
    ReadDeepScanReportTarget = ""

    Dim raw As String
    ' Use the same reader as footer HTML (more reliable than the thin helper).
    raw = ReadTextUtf8(outputPath)
    If Len(raw) = 0 Then raw = ReadTextFileUtf8(outputPath)
    If Len(raw) = 0 Then
        MSCANModLogging.WriteLog "ReadDeepScanReportTarget: empty read from " & outputPath
        Exit Function
    End If

    ' Strip UTF-8 BOM / NUL noise that breaks Left$("report_url=") matching.
    raw = Replace(raw, Chr$(0), "")
    If Left$(raw, 1) = ChrW(&HFEFF) Then raw = Mid$(raw, 2)
    If Left$(raw, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then raw = Mid$(raw, 4)

    Dim line As String
    Dim lines() As String
    Dim i As Long
    Dim p As Long
    lines = Split(Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Then GoTo ContinueDeepLine

        p = InStr(1, line, "report_url=", vbTextCompare)
        If p > 0 Then
            ReadDeepScanReportTarget = Trim$(Mid$(line, p + Len("report_url=")))
            Exit Function
        End If
        p = InStr(1, line, "report_path=", vbTextCompare)
        If p > 0 Then
            ReadDeepScanReportTarget = Trim$(Mid$(line, p + Len("report_path=")))
            Exit Function
        End If
        ' Bare file:// or Windows path
        If InStr(1, line, "file:", vbTextCompare) > 0 Then
            p = InStr(1, line, "file:", vbTextCompare)
            ReadDeepScanReportTarget = Trim$(Mid$(line, p))
            Exit Function
        End If
        If (Len(line) >= 3 And Mid$(line, 2, 1) = ":" And (Mid$(line, 3, 1) = "\" Or Mid$(line, 3, 1) = "/")) Then
            ReadDeepScanReportTarget = line
            Exit Function
        End If
ContinueDeepLine:
    Next i

    MSCANModLogging.WriteLog "ReadDeepScanReportTarget: no report_url in " & outputPath & _
        " (len=" & Len(raw) & " head=" & Left$(Replace(Replace(raw, vbCrLf, " "), vbLf, " "), 80) & ")"
    Exit Function
EH:
    MSCANModLogging.WriteLog "ReadDeepScanReportTarget error: #" & Err.Number & " - " & Err.Description & " path=" & outputPath
    ReadDeepScanReportTarget = ""
End Function

Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    On Error GoTo EH
    ReadTextFileUtf8 = ""
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadTextFileUtf8 = stream.ReadText(-1)
    stream.Close
    Exit Function
EH:
    On Error Resume Next
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        Set ts = fso.OpenTextFile(filePath, 1)
        ReadTextFileUtf8 = ts.ReadAll
        ts.Close
    End If
End Function

' Last-resort: newest deepscan_*.html under output\deepscan_reports.
Private Function FindLatestDeepScanReport() As String
    On Error GoTo EH
    FindLatestDeepScanReport = ""
    Dim fso As Object
    Dim folder As Object
    Dim f As Object
    Dim bestPath As String
    Dim bestDate As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = "C:\GeoFooter\output\deepscan_reports"
    If Not fso.FolderExists(dir) Then Exit Function
    Set folder = fso.GetFolder(dir)
    bestDate = 0
    For Each f In folder.Files
        If LCase$(fso.GetExtensionName(f.Name)) = "html" Then
            If InStr(1, LCase$(f.Name), "deepscan_", vbTextCompare) = 1 Then
                If f.DateLastModified > bestDate Then
                    bestDate = f.DateLastModified
                    bestPath = f.Path
                End If
            End If
        End If
    Next f
    FindLatestDeepScanReport = bestPath
    Exit Function
EH:
    FindLatestDeepScanReport = ""
End Function

Private Sub OpenDeepScanReport(ByVal target As String)
    On Error GoTo EH
    If Len(target) = 0 Then Exit Sub

    Dim pathOrUrl As String
    pathOrUrl = Trim$(target)

    ' Prefer opening local path via Shell when file:// is present
    If InStr(1, pathOrUrl, "file:///", vbTextCompare) = 1 Then
        pathOrUrl = Mid$(pathOrUrl, 9)
        pathOrUrl = Replace(pathOrUrl, "/", "\")
        ' file:///C:/... → C:\...
    ElseIf InStr(1, pathOrUrl, "file://", vbTextCompare) = 1 Then
        pathOrUrl = Mid$(pathOrUrl, 8)
        pathOrUrl = Replace(pathOrUrl, "/", "\")
    End If

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run """" & pathOrUrl & """", 1, False
    Exit Sub
EH:
    MSCANModLogging.WriteLog "OpenDeepScanReport error: #" & Err.Number & " - " & Err.Description & " target=" & target
End Sub

' Read-only header export for Python. Does not modify or save the mail item.
' deepMode=True also exports body HTML/text sidecars for Deep Scan analyzers.
Private Function ExportHeadersForGeolocation(ByVal mail As Object, Optional ByVal deepMode As Boolean = False) As String
    On Error GoTo EH
    ExportHeadersForGeolocation = ""

    MSCANModLogging.WriteLog "ExportHeadersForGeolocation: Starting for subject: " & SafeSubject(mail) & " deep=" & CStr(deepMode)

    ' --- Define Headers to Strip ---
    Dim headersToStrip As Object
    Set headersToStrip = CreateObject("Scripting.Dictionary")
    headersToStrip.CompareMode = vbTextCompare ' Case-insensitive

    headersToStrip.Add "Received", True
    headersToStrip.Add "Received-SPF", True
    headersToStrip.Add "X-Originating-IP", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-OriginalArrivalTime", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-Network-Message-Id", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-Id", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-FromEntityHeader", True
    headersToStrip.Add "X-MS-Exchange-Transport-CrossTenantHeadersStamped", True

    ' --- Read Headers ---
    Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
    Dim pa As Outlook.PropertyAccessor
    Set pa = mail.PropertyAccessor

    Dim currentHeaders As String
    On Error Resume Next
    currentHeaders = pa.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    If Err.Number <> 0 Then
        MSCANModLogging.WriteLog "ExportHeadersForGeolocation: Could not read PR_TRANSPORT_MESSAGE_HEADERS for subject: " & SafeSubject(mail) & " Error: " & Err.Description
        Err.Clear
        Set pa = Nothing
        Exit Function
    End If
    On Error GoTo EH

    If Len(currentHeaders) = 0 Then
        MSCANModLogging.WriteLog "ExportHeadersForGeolocation: No headers found for subject: " & SafeSubject(mail)
        Set pa = Nothing
        Exit Function
    End If

    Dim headerLines() As String
    headerLines = Split(currentHeaders, vbCrLf)

    Dim newHeader As String
    Dim removedHeadersLog As String
    newHeader = ""
    removedHeadersLog = ""

    Dim isStripping As Boolean: isStripping = False
    Dim i As Long
    Dim headerLine As String
    Dim headerKey As String
    Dim pos As Integer

    For i = 0 To UBound(headerLines)
        headerLine = headerLines(i)
        pos = InStr(headerLine, ":")

        If pos > 0 Then
            ' New header key-value pair
            headerKey = Trim(left(headerLine, pos - 1))
            If headersToStrip.Exists(headerKey) Then
                isStripping = True
                removedHeadersLog = removedHeadersLog & headerLine & vbCrLf
            Else
                isStripping = False
                newHeader = newHeader & headerLine & vbCrLf
            End If
        Else
            ' Continuation line (no key)
            If isStripping Then
                removedHeadersLog = removedHeadersLog & headerLine & vbCrLf
            Else
                newHeader = newHeader & headerLine & vbCrLf
            End If
        End If
    Next i

    If Len(removedHeadersLog) > 0 Then
        MSCANModLogging.WriteLog "ExportHeadersForGeolocation: Sensitive headers will be stripped after footer insert (" & _
            Len(removedHeadersLog) & " chars)."
    Else
        MSCANModLogging.WriteLog "ExportHeadersForGeolocation: No strip targets found in transport headers."
    End If

    ' Always export FULL transport headers for Python (From, Auth-Results, etc.)
    Dim fileContent As String
    Dim savedFilePath As String
    Dim attachmentSummary As String
    Dim beaconSummary As String
    Dim attachmentScanDir As String
    Dim deepPreamble As String
    deepPreamble = ""

    ' Per-sender attachment blocking (footer "Block attachments from sender"
    ' button) — quarantine + remove before export so nothing gets scanned or
    ' left on the mail.
    Dim quarantinedCount As Long
    quarantinedCount = 0
    If MSCANModSenderRules.ShouldBlockAttachments(mail) Then
        quarantinedCount = MSCANModSenderRules.QuarantineAttachments(mail)
    End If

    m_LastAttachmentScanDir = ExportAttachmentsForScan(mail)
    attachmentScanDir = m_LastAttachmentScanDir
    attachmentSummary = BuildAttachmentSummary(mail)
    If quarantinedCount > 0 Then
        attachmentSummary = attachmentSummary & "; quarantined=" & quarantinedCount
    End If
    beaconSummary = BuildBeaconSummary(mail)
    m_LastDeepPayloadDir = ""

    ' Body is exported for every scan so links can be safety-checked in Basic
    ' scans too (Deep additionally runs the body-content scanner on it).
    m_LastDeepPayloadDir = ExportDeepBodyPayload(mail)
    If deepMode Then deepPreamble = "AES-Scan-Mode: deep" & vbCrLf
    If Len(m_LastDeepPayloadDir) > 0 Then
        deepPreamble = deepPreamble & _
            "AES-Body-Dir: " & m_LastDeepPayloadDir & vbCrLf & _
            "AES-Body-Html-File: " & m_LastDeepPayloadDir & "\body.html" & vbCrLf & _
            "AES-Body-Text-File: " & m_LastDeepPayloadDir & "\body.txt" & vbCrLf
    End If

    fileContent = "Original Subject: " & SafeSubject(mail) & vbCrLf & _
                  "AES-Attachments: " & attachmentSummary & vbCrLf & _
                  "AES-Attachments-Dir: " & attachmentScanDir & vbCrLf & _
                  "AES-Beacons: " & beaconSummary & vbCrLf & _
                  deepPreamble & _
                  "Processing Timestamp: " & Now() & vbCrLf & _
                  "----------------------------------------" & vbCrLf & _
                  currentHeaders

    savedFilePath = SaveHeadersToFile(fileContent)

    If savedFilePath <> "" Then
        MSCANModLogging.WriteLog "ExportHeadersForGeolocation: Header file saved: " & savedFilePath
        ExportHeadersForGeolocation = savedFilePath
    Else
        MSCANModLogging.WriteLog "ExportHeadersForGeolocation: CRITICAL - FAILED to save headers file."
    End If

    Set pa = Nothing
    Set headersToStrip = Nothing
    Exit Function

EH:
    MSCANModLogging.WriteLog "ExportHeadersForGeolocation FATAL error: #" & Err.Number & " - " & Err.Description & " for subject: " & SafeSubject(mail)
End Function

Private Function ExportDeepBodyPayload(ByVal mail As Object) As String
    On Error GoTo EH
    ExportDeepBodyPayload = ""

    Dim payloadDir As String
    payloadDir = GetBaseDir() & "\deep_payload\scan_" & Format(Now, "yyyymmdd_hhnnss") & "_" & _
                 Format(Int((Timer - Int(Timer)) * 1000), "000")

    If Not EnsureFolderExists(payloadDir) Then
        MSCANModLogging.WriteLog "ExportDeepBodyPayload: Could not create folder: " & payloadDir
        Exit Function
    End If

    Dim htmlBody As String
    Dim textBody As String
    On Error Resume Next
    htmlBody = ""
    textBody = ""
    htmlBody = mail.HTMLBody
    Err.Clear
    textBody = mail.Body
    On Error GoTo EH
    If Len(htmlBody) = 0 And Len(textBody) > 0 Then htmlBody = textBody

    WriteTextUtf8 payloadDir & "\body.html", htmlBody
    WriteTextUtf8 payloadDir & "\body.txt", textBody

    MSCANModLogging.WriteLog "ExportDeepBodyPayload: Wrote body sidecars to " & payloadDir
    ExportDeepBodyPayload = payloadDir
    Exit Function

EH:
    MSCANModLogging.WriteLog "ExportDeepBodyPayload error: #" & Err.Number & " - " & Err.Description
End Function

Private Function BuildAttachmentSummary(ByVal mail As Object) As String
    On Error Resume Next
    BuildAttachmentSummary = "count=" & mail.Attachments.Count & "; ok=0; not_ok=0"
End Function

Private Function ExportAttachmentsForScan(ByVal mail As Object) As String
    On Error GoTo EH
    ExportAttachmentsForScan = ""

    If mail.Attachments.Count = 0 Then Exit Function

    Dim scanDir As String
    scanDir = GetBaseDir() & "\attachments\scan_" & Format(Now, "yyyymmdd_hhnnss") & "_" & _
              Format(Int((Timer - Int(Timer)) * 1000), "000")

    If Not EnsureFolderExists(scanDir) Then
        MSCANModLogging.WriteLog "ExportAttachmentsForScan: Could not create folder: " & scanDir
        Exit Function
    End If

    Dim i As Long
    Dim savedCount As Long
    savedCount = 0

    For i = 1 To mail.Attachments.Count
        Dim att As Outlook.Attachment
        Set att = mail.Attachments(i)

        Dim safeName As String
        safeName = SanitizeAttachmentFileName(att.FileName)
        If Len(safeName) = 0 Then safeName = "attachment_" & CStr(i)

        Dim targetPath As String
        targetPath = scanDir & "\" & safeName
        targetPath = UniqueAttachmentPath(targetPath)

        On Error Resume Next
        att.SaveAsFile targetPath
        If Err.Number = 0 Then
            savedCount = savedCount + 1
        Else
            MSCANModLogging.WriteLog "ExportAttachmentsForScan: Failed to save " & att.FileName & " — " & Err.Description
            Err.Clear
        End If
        On Error GoTo EH
    Next i

    If savedCount = 0 Then
        CleanupAttachmentScanDir scanDir
        Exit Function
    End If

    MSCANModLogging.WriteLog "ExportAttachmentsForScan: Saved " & savedCount & " attachment(s) to " & scanDir
    ExportAttachmentsForScan = scanDir
    Exit Function

EH:
    MSCANModLogging.WriteLog "ExportAttachmentsForScan error: #" & Err.Number & " - " & Err.Description
End Function

Private Function SanitizeAttachmentFileName(ByVal fileName As String) As String
    Dim cleaned As String
    Dim i As Long
    Dim ch As String

    cleaned = fileName
    cleaned = Replace(cleaned, "/", "_")
    cleaned = Replace(cleaned, "\", "_")
    cleaned = Replace(cleaned, ":", "_")
    cleaned = Replace(cleaned, "*", "_")
    cleaned = Replace(cleaned, "?", "_")
    cleaned = Replace(cleaned, """", "_")
    cleaned = Replace(cleaned, "<", "_")
    cleaned = Replace(cleaned, ">", "_")
    cleaned = Replace(cleaned, "|", "_")

    If Len(cleaned) > 120 Then cleaned = Right$(cleaned, 120)
    SanitizeAttachmentFileName = cleaned
End Function

Private Function UniqueAttachmentPath(ByVal targetPath As String) As String
    Dim fso As Object
    Dim candidate As String
    Dim seq As Long
    Dim dotPos As Long
    Dim baseName As String
    Dim extPart As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    candidate = targetPath
    seq = 1
    dotPos = InStrRev(targetPath, ".")

    Do While fso.FileExists(candidate)
        If dotPos > 0 Then
            baseName = Left$(targetPath, dotPos - 1)
            extPart = Mid$(targetPath, dotPos)
            candidate = baseName & "_" & CStr(seq) & extPart
        Else
            candidate = targetPath & "_" & CStr(seq)
        End If
        seq = seq + 1
    Loop

    UniqueAttachmentPath = candidate
    Set fso = Nothing
End Function

Public Sub CleanupAttachmentScanDir(ByVal scanDir As String)
    On Error Resume Next
    If Len(scanDir) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(scanDir) Then
        fso.DeleteFolder scanDir, True
        MSCANModLogging.WriteLog "CleanupAttachmentScanDir: Removed " & scanDir
    End If
    Set fso = Nothing
End Sub

Private Function AttachmentsAreOk(ByVal mail As Object) As Boolean
    On Error Resume Next
    Dim i As Long
    For i = 1 To mail.Attachments.Count
        If IsRiskyAttachmentName(mail.Attachments(i).FileName) Then
            AttachmentsAreOk = False
            Exit Function
        End If
    Next i
    AttachmentsAreOk = True
End Function

Private Function IsRiskyAttachmentName(ByVal fileName As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))
    Select Case ext
        Case "exe", "bat", "cmd", "com", "scr", "pif", "vbs", "vbe", "js", "jse", "wsf", "wsh", "ps1", "msi", "msp", "hta", "dll", "lnk", "iso", "img", "jar", "reg", "inf"
            IsRiskyAttachmentName = True
        Case "docm", "xlsm", "pptm", "dotm", "xltm", "potm", "ppam", "sldm"
            IsRiskyAttachmentName = True
        Case Else
            IsRiskyAttachmentName = False
    End Select
End Function

Private Function BuildBeaconSummary(ByVal mail As Object) As String
    On Error Resume Next
    Dim beaconCount As Long
    Dim beaconUrls As String
    Dim blockedCount As Long

    beaconUrls = ""
    beaconCount = 0
    blockedCount = 0

    If mail.BodyFormat = olFormatHTML Then
        ' Detect on the original body first, then neutralise if rules say so,
        ' so the footer can report "X detected / Y blocked".
        beaconCount = CollectEmailBeacons(mail.HTMLBody, beaconUrls)
        If beaconCount > 0 Then
            If MSCANModSenderRules.ShouldBlockBeacons(mail) Then
                blockedCount = MSCANModSenderRules.NeutralizeBeaconsInMail(mail)
            End If
        End If
    End If

    If beaconCount = 0 Then
        BuildBeaconSummary = "count=0; status=None"
    Else
        ' Keep blocked= before urls= — Python's urls regex consumes to end of line.
        BuildBeaconSummary = "count=" & beaconCount & "; status=Detected; blocked=" & _
            blockedCount & "; urls=" & beaconUrls
    End If
End Function

Private Function CollectEmailBeacons(ByVal htmlBody As String, ByRef beaconUrls As String) As Long
    On Error Resume Next
    Dim count As Long
    Dim pos As Long
    Dim tagStart As Long
    Dim tagEnd As Long
    Dim tagText As String
    Dim urlList As String
    Dim imgUrl As String

    count = 0
    urlList = ""
    pos = 1

    Do
        tagStart = InStr(pos, htmlBody, "<img", vbTextCompare)
        If tagStart = 0 Then Exit Do
        tagEnd = InStr(tagStart, htmlBody, ">")
        If tagEnd = 0 Then Exit Do
        tagText = Mid$(htmlBody, tagStart, tagEnd - tagStart + 1)
        If IsLikelyTrackingBeacon(tagText) Then
            count = count + 1
            imgUrl = ExtractImgSrc(tagText)
            If Len(imgUrl) > 0 Then
                If Len(urlList) > 0 Then urlList = urlList & "|"
                urlList = urlList & imgUrl
            End If
        End If
        pos = tagEnd + 1
    Loop

    pos = 1
    Do
        tagStart = InStr(pos, htmlBody, "<v:imagedata", vbTextCompare)
        If tagStart = 0 Then Exit Do
        tagEnd = InStr(tagStart, htmlBody, ">")
        If tagEnd = 0 Then Exit Do
        tagText = Mid$(htmlBody, tagStart, tagEnd - tagStart + 1)
        If IsLikelyTrackingBeacon(tagText) Then
            count = count + 1
            imgUrl = ExtractVmlImageSrc(tagText)
            If Len(imgUrl) > 0 Then
                If Len(urlList) > 0 Then urlList = urlList & "|"
                urlList = urlList & imgUrl
            End If
        End If
        pos = tagEnd + 1
    Loop

    If Len(urlList) > 1800 Then urlList = Left$(urlList, 1800)
    beaconUrls = urlList
    CollectEmailBeacons = count
End Function

Private Function ExtractImgSrc(ByVal imgTag As String) As String
    Dim p As Long
    Dim q As Long
    Dim ch As String

    p = InStr(1, imgTag, "src=", vbTextCompare)
    If p = 0 Then Exit Function

    p = p + 4
    ch = Mid$(imgTag, p, 1)
    If ch = """" Or ch = "'" Then
        q = InStr(p + 1, imgTag, ch)
        If q > p + 1 Then ExtractImgSrc = Mid$(imgTag, p + 1, q - p - 1)
    Else
        q = InStr(p, imgTag, " ")
        If q = 0 Then q = InStr(p, imgTag, ">")
        If q > p Then ExtractImgSrc = Mid$(imgTag, p, q - p)
    End If
End Function

Private Function ExtractVmlImageSrc(ByVal tagText As String) As String
    Dim p As Long
    Dim q As Long

    p = InStr(1, tagText, "src=", vbTextCompare)
    If p = 0 Then Exit Function

    p = p + 4
    If Mid$(tagText, p, 1) = """" Then
        q = InStr(p + 1, tagText, """")
        If q > p + 1 Then ExtractVmlImageSrc = Mid$(tagText, p + 1, q - p - 1)
    End If
End Function

Public Function IsLikelyTrackingBeacon(ByVal imgTag As String) As Boolean
    Dim t As String
    t = LCase$(imgTag)

    If InStr(t, "width=""1""") > 0 Or InStr(t, "width='1'") > 0 Or InStr(t, "width=1") > 0 Then
        IsLikelyTrackingBeacon = True
        Exit Function
    End If
    If InStr(t, "height=""1""") > 0 Or InStr(t, "height='1'") > 0 Or InStr(t, "height=1") > 0 Then
        IsLikelyTrackingBeacon = True
        Exit Function
    End If
    If InStr(t, "display:none") > 0 Or InStr(t, "visibility:hidden") > 0 Or InStr(t, "opacity:0") > 0 Then
        IsLikelyTrackingBeacon = True
        Exit Function
    End If

    If InStr(t, "mailtrack") > 0 Or InStr(t, "/track") > 0 Or InStr(t, "/open?") > 0 Or InStr(t, "pixel") > 0 Then
        IsLikelyTrackingBeacon = True
        Exit Function
    End If
    If InStr(t, "sendgrid.net/wf") > 0 Or InStr(t, "list-manage.com") > 0 Or InStr(t, "hubspot.com") > 0 Then
        IsLikelyTrackingBeacon = True
        Exit Function
    End If

    IsLikelyTrackingBeacon = False
End Function

Private Sub StripSensitiveHeadersFromMail(ByVal mail As Object)
    On Error GoTo EH

    Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
    Dim pa As Outlook.PropertyAccessor
    Dim currentHeaders As String
    Dim headersToStrip As Object
    Dim headerLines() As String
    Dim newHeader As String
    Dim removedHeadersLog As String
    Dim isStripping As Boolean
    Dim i As Long
    Dim headerLine As String
    Dim headerKey As String
    Dim pos As Integer

    Set pa = mail.PropertyAccessor
    On Error Resume Next
    currentHeaders = pa.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    If Err.Number <> 0 Or Len(currentHeaders) = 0 Then Exit Sub
    On Error GoTo EH

    Set headersToStrip = CreateObject("Scripting.Dictionary")
    headersToStrip.CompareMode = vbTextCompare
    headersToStrip.Add "Received", True
    headersToStrip.Add "Received-SPF", True
    headersToStrip.Add "X-Originating-IP", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-OriginalArrivalTime", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-Network-Message-Id", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-Id", True
    headersToStrip.Add "X-MS-Exchange-CrossTenant-FromEntityHeader", True
    headersToStrip.Add "X-MS-Exchange-Transport-CrossTenantHeadersStamped", True

    headerLines = Split(currentHeaders, vbCrLf)
    newHeader = ""
    removedHeadersLog = ""
    isStripping = False

    For i = 0 To UBound(headerLines)
        headerLine = headerLines(i)
        pos = InStr(headerLine, ":")
        If pos > 0 Then
            headerKey = Trim$(Left$(headerLine, pos - 1))
            If headersToStrip.Exists(headerKey) Then
                isStripping = True
                removedHeadersLog = removedHeadersLog & headerLine & vbCrLf
            Else
                isStripping = False
                newHeader = newHeader & headerLine & vbCrLf
            End If
        Else
            If isStripping Then
                removedHeadersLog = removedHeadersLog & headerLine & vbCrLf
            Else
                newHeader = newHeader & headerLine & vbCrLf
            End If
        End If
    Next i

    If Len(removedHeadersLog) = 0 Then Exit Sub

    pa.SetProperty PR_TRANSPORT_MESSAGE_HEADERS, newHeader
    mail.Save
    MSCANModLogging.WriteLog "StripSensitiveHeadersFromMail: Sensitive headers removed after footer insert."
    Exit Sub

EH:
    MSCANModLogging.WriteLog "StripSensitiveHeadersFromMail error: #" & Err.Number & " - " & Err.Description
End Sub

Private Function WaitForMailReady(ByVal mail As Object) As Boolean
    On Error GoTo EH

    Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
    Dim attempt As Long

    ' Keep this short — WaitForMailReady runs on the Outlook UI thread.
    ' Prefer requeue (ProcessEmail) over multi-second busy-waits.
    For attempt = 1 To 3
        On Error Resume Next
        Dim entryId As String
        Dim headers As String

        entryId = mail.EntryID
        Err.Clear
        headers = mail.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
        If Err.Number <> 0 Or Len(headers) = 0 Then
            Err.Clear
            headers = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
        End If
        If Err.Number = 0 And Len(entryId) > 0 And Len(headers) > 0 Then
            WaitForMailReady = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo EH

        If attempt < 3 Then PauseSeconds 0.2
    Next attempt

    WaitForMailReady = False
    Exit Function

EH:
    WaitForMailReady = False
End Function

Private Sub PauseSeconds(ByVal sec As Single)
    Dim t As Single
    t = Timer
    Do While Timer < t + sec
        DoEvents
    Loop
End Sub

'===============================================================================
' PRIVATE: Saves header content to a timestamped file.
' Returns the full path of the saved file on success, or an empty string on failure.
'===============================================================================
Private Function SaveHeadersToFile(ByVal headerContent As String) As String
    On Error GoTo EH

    Dim baseDir As String
    baseDir = GetBaseDir()
    If Len(baseDir) = 0 Then
        MSCANModLogging.WriteLog "SaveHeadersToFile: No base directory available."
        SaveHeadersToFile = ""
        Exit Function
    End If

    Dim headerSavePath As String
    headerSavePath = baseDir
    If Right(headerSavePath, 1) <> "\" Then headerSavePath = headerSavePath & "\"
    headerSavePath = headerSavePath & "headers\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(headerSavePath) Then
        On Error Resume Next
        fso.CreateFolder headerSavePath
        If Err.Number <> 0 Then
            MSCANModLogging.WriteLog "SaveHeadersToFile: Could not create folder: " & headerSavePath & " Error: " & Err.Description
            Err.Clear
            SaveHeadersToFile = ""
            Exit Function
        End If
        On Error GoTo EH
        MSCANModLogging.WriteLog "SaveHeadersToFile: Created directory: " & headerSavePath
    End If

    ' --- Generate unique filename with milliseconds + process id fallback ---
    Dim milliseconds As String
    milliseconds = Format(Int((Timer - Int(Timer)) * 1000), "000")

    Dim fileName As String
    fileName = "headers_" & Format(Now, "yyyymmdd_hhnnss") & "_" & milliseconds & ".txt"
    Dim fullPath As String
    fullPath = headerSavePath & fileName

    ' --- Write the file (UTF-8) ---
    WriteTextUtf8 fullPath, headerContent

    SaveHeadersToFile = fullPath

    Set fso = Nothing
    Exit Function

EH:
    MSCANModLogging.WriteLog "SaveHeadersToFile FAILED: Error #" & Err.Number & " - " & Err.Description
    SaveHeadersToFile = ""
    If Not fso Is Nothing Then Set fso = Nothing
End Function


'===============================================================================
' ASYNC WORKER: start Python outside Outlook's UI thread, then callback when done.
' WScript waits on Python; Outlook only spends a brief moment applying the footer.
'===============================================================================
Private Function StartAsyncGeolocationJob(ByVal mail As Object, ByVal headerFilePath As String, ByVal footerMode As String) As Boolean
    On Error GoTo ErrHandler
    StartAsyncGeolocationJob = False

    If Len(footerMode) = 0 Then footerMode = "compact"
    footerMode = LCase$(footerMode)
    If footerMode <> "full" And footerMode <> "deep" Then footerMode = "compact"

    Dim pythonExe As String: pythonExe = GetPythonExe()
    Dim pythonScript As String: pythonScript = GetPythonScript()
    If Len(pythonExe) = 0 Or Len(pythonScript) = 0 Then
        MSCANModLogging.WriteLog "StartAsyncGeolocationJob: Missing Python executable or script path."
        Exit Function
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(headerFilePath) Then
        MSCANModLogging.WriteLog "StartAsyncGeolocationJob: Header file does not exist: " & headerFilePath
        Exit Function
    End If

    Dim entryId As String
    entryId = mail.EntryID
    If Len(entryId) = 0 Then
        MSCANModLogging.WriteLog "StartAsyncGeolocationJob: Mail has no EntryID yet."
        Exit Function
    End If

    Dim workDir As String
    workDir = fso.GetParentFolderName(pythonScript)
    If LCase$(fso.GetFileName(workDir)) = "vba" Then workDir = fso.GetParentFolderName(workDir)

    Dim outputDir As String
    outputDir = workDir & "\output"
    If Not EnsureFolderExists(outputDir) Then
        outputDir = GetBaseDir() & "\output"
        EnsureFolderExists outputDir
    End If

    m_AsyncJobSeq = m_AsyncJobSeq + 1
    Dim jobId As String
    jobId = "aes" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(m_AsyncJobSeq)

    Dim outputFile As String
    If footerMode = "deep" Then
        outputFile = outputDir & "\deep_" & jobId & ".txt"
    Else
        outputFile = outputDir & "\geo_" & jobId & ".html"
    End If
    ClearJobSidecars outputFile

    Dim runExe As String
    runExe = PreferPythonW(pythonExe)

    Dim cmd As String
    cmd = """" & runExe & """ """ & pythonScript & """ """ & headerFilePath & """ """ & outputFile & """ " & footerMode

    EnsureAsyncJobs
    Dim job As Object
    Set job = CreateObject("Scripting.Dictionary")
    job.CompareMode = vbTextCompare
    job("EntryID") = entryId
    job("FooterPath") = outputFile
    job("Mode") = footerMode
    job("AttachDir") = m_LastAttachmentScanDir
    job("PayloadDir") = m_LastDeepPayloadDir
    job("Subject") = SafeSubject(mail)
    job("StartedAt") = Now
    ' Must use Set when storing an Object in a Dictionary (else runtime error 450).
    Set m_AsyncJobs(jobId) = job
    ' Ownership transferred to the async job; do not clean up here on success.
    m_LastAttachmentScanDir = ""
    m_LastDeepPayloadDir = ""

    Dim vbsPath As String
    vbsPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_geo_job_" & jobId & ".vbs"
    If Not WriteAsyncGeoJobScript(vbsPath, workDir, cmd, jobId, outputFile) Then
        MSCANModLogging.WriteLog "StartAsyncGeolocationJob: could not write job script."
        m_AsyncJobs.Remove jobId
        NotifyScanBusyUi
        Exit Function
    End If

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    ' WaitOnReturn=False — Outlook returns immediately; WScript blocks on Python instead.
    shell.Run "wscript.exe //B //Nologo """ & vbsPath & """", 0, False

    MSCANModLogging.WriteLog "StartAsyncGeolocationJob: launched job " & jobId & " mode=" & footerMode & " subject=" & SafeSubject(mail)
    MSCANModLogging.WriteLog "StartAsyncGeolocationJob: cmd=" & cmd
    StartAsyncGeolocationJob = True
    NotifyScanBusyUi
    Exit Function

ErrHandler:
    MSCANModLogging.WriteLog "StartAsyncGeolocationJob error: #" & Err.Number & " - " & Err.Description
    StartAsyncGeolocationJob = False
End Function

' This Outlook build does not expose ThisOutlookSession publics on the COM
' Application object (and Outlook has no Application.Run), so the script cannot
' call back into VBA directly. Instead: on failure it drops a "<output>.fail"
' marker, then "nudges" Outlook by reading inbox items — that fires
' Application_ItemLoad inside Outlook, whose handler runs NudgeAsyncWork, which
' reconciles pending jobs (applies the footer / clears the busy icon).
'
' Critical: wait on the output file, NOT pythonw process exit. pythonw often
' hangs on interpreter shutdown after the HTML is already written, which used
' to leave AES PROC stuck for minutes. After the file is ready, keep nudging
' until VBA writes "<output>.applied" (or we time out).
Private Function WriteAsyncGeoJobScript(ByVal vbsPath As String, ByVal workDir As String, ByVal commandLine As String, ByVal jobId As String, ByVal outputFile As String) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Dim ts As Object
    Dim folder As String
    Dim logPath As String
    Dim appliedPath As String
    Dim failPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(vbsPath)
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder

    logPath = Environ$("LOCALAPPDATA") & "\GeoFooter\Logs\aes_geo_callback.log"
    appliedPath = outputFile & ".applied"
    failPath = outputFile & ".fail"

    Set ts = fso.CreateTextFile(vbsPath, True, False)
    ts.WriteLine "Dim sh, code, prev, ol, f, fso, itm, items, i, n, dummy"
    ts.WriteLine "On Error Resume Next"
    ts.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    ts.WriteLine "Set sh = CreateObject(""WScript.Shell"")"
    ts.WriteLine "prev = sh.CurrentDirectory"
    ts.WriteLine "sh.CurrentDirectory = " & VbsQuoteString(workDir)
    ts.WriteLine "Err.Clear"
    ' Do not WaitOnReturn — pythonw can hang after writing the report.
    ts.WriteLine "code = sh.Run(" & VbsQuoteString(commandLine) & ", 0, False)"
    ts.WriteLine "If Err.Number <> 0 Then code = 99"
    ts.WriteLine "sh.CurrentDirectory = prev"
    ts.WriteLine "' Poll until report file has content (up to ~5 minutes)."
    ts.WriteLine "code = 1"
    ts.WriteLine "For i = 1 To 600"
    ts.WriteLine "  If fso.FileExists(" & VbsQuoteString(outputFile) & ") Then"
    ts.WriteLine "    If fso.GetFile(" & VbsQuoteString(outputFile) & ").Size > 0 Then"
    ts.WriteLine "      code = 0"
    ts.WriteLine "      Exit For"
    ts.WriteLine "    End If"
    ts.WriteLine "  End If"
    ts.WriteLine "  If fso.FileExists(" & VbsQuoteString(failPath) & ") Then"
    ts.WriteLine "    code = 1"
    ts.WriteLine "    Exit For"
    ts.WriteLine "  End If"
    ts.WriteLine "  WScript.Sleep 500"
    ts.WriteLine "Next"
    ts.WriteLine "If code <> 0 And Not fso.FileExists(" & VbsQuoteString(outputFile) & ") Then"
    ts.WriteLine "  If Not fso.FileExists(" & VbsQuoteString(failPath) & ") Then"
    ts.WriteLine "    Set f = fso.CreateTextFile(" & VbsQuoteString(failPath) & ", True)"
    ts.WriteLine "    f.WriteLine ""timeout_or_exit=1"""
    ts.WriteLine "    f.Close"
    ts.WriteLine "  End If"
    ts.WriteLine "End If"
    ts.WriteLine "Set ol = GetObject(, ""Outlook.Application"")"
    ts.WriteLine "If ol Is Nothing Then"
    ts.WriteLine "  Set f = fso.OpenTextFile(" & VbsQuoteString(logPath) & ", 8, True)"
    ts.WriteLine "  f.WriteLine Now & "" | GetObject Outlook failed for job " & jobId & """"
    ts.WriteLine "  f.Close"
    ts.WriteLine "  WScript.Quit 2"
    ts.WriteLine "End If"
    ts.WriteLine "' Rotate inbox items so ItemLoad fires even when GetFirst is cached."
    ts.WriteLine "Set f = fso.OpenTextFile(" & VbsQuoteString(logPath) & ", 8, True)"
    ts.WriteLine "f.WriteLine Now & "" | nudge start job=" & jobId & " code="" & code"
    ts.WriteLine "f.Close"
    ts.WriteLine "For n = 1 To 45"
    ts.WriteLine "  Err.Clear"
    ts.WriteLine "  Set items = Nothing"
    ts.WriteLine "  Set itm = Nothing"
    ts.WriteLine "  Set items = ol.GetNamespace(""MAPI"").GetDefaultFolder(6).Items"
    ts.WriteLine "  If items Is Nothing Or Err.Number <> 0 Then"
    ts.WriteLine "    Err.Clear"
    ts.WriteLine "    Set items = ol.GetNamespace(""MAPI"").GetDefaultFolder(5).Items"
    ts.WriteLine "  End If"
    ts.WriteLine "  If Not items Is Nothing Then"
    ts.WriteLine "    items.Sort ""[ReceivedTime]"", True"
    ts.WriteLine "    If (n Mod 2) = 1 Then"
    ts.WriteLine "      Set itm = items.GetLast"
    ts.WriteLine "    Else"
    ts.WriteLine "      Set itm = items.GetFirst"
    ts.WriteLine "    End If"
    ts.WriteLine "    If Not itm Is Nothing Then dummy = itm.EntryID"
    ts.WriteLine "    If (n Mod 3) = 0 Then"
    ts.WriteLine "      Set itm = items.GetNext"
    ts.WriteLine "      If Not itm Is Nothing Then dummy = itm.Subject"
    ts.WriteLine "    End If"
    ts.WriteLine "  End If"
    ts.WriteLine "  If fso.FileExists(" & VbsQuoteString(appliedPath) & ") Then"
    ts.WriteLine "    Set f = fso.OpenTextFile(" & VbsQuoteString(logPath) & ", 8, True)"
    ts.WriteLine "    f.WriteLine Now & "" | nudge applied job=" & jobId & " after="" & n"
    ts.WriteLine "    f.Close"
    ts.WriteLine "    WScript.Quit code"
    ts.WriteLine "  End If"
    ts.WriteLine "  WScript.Sleep 2000"
    ts.WriteLine "Next"
    ts.WriteLine "Set f = fso.OpenTextFile(" & VbsQuoteString(logPath) & ", 8, True)"
    ts.WriteLine "f.WriteLine Now & "" | nudge timeout job=" & jobId & " code="" & code"
    ts.WriteLine "f.Close"
    ts.WriteLine "WScript.Quit code"
    ts.Close
    WriteAsyncGeoJobScript = True
    Exit Function

EH:
    WriteAsyncGeoJobScript = False
End Function

Private Sub EnsureAsyncJobs()
    If m_AsyncJobs Is Nothing Then
        Set m_AsyncJobs = CreateObject("Scripting.Dictionary")
        m_AsyncJobs.CompareMode = vbTextCompare
    End If
End Sub

Private Function ResolveMailByEntryID(ByVal entryId As String) As Object
    On Error GoTo EH
    Set ResolveMailByEntryID = Nothing
    If Len(entryId) = 0 Then Exit Function
    Dim itm As Object
    Set itm = Application.Session.GetItemFromID(entryId)
    If IsScannableItem(itm) Then Set ResolveMailByEntryID = itm
    Exit Function
EH:
    Set ResolveMailByEntryID = Nothing
End Function

Private Function PreferPythonW(ByVal pythonExe As String) As String
    On Error Resume Next
    PreferPythonW = pythonExe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If LCase$(fso.GetFileName(pythonExe)) = "python.exe" Then
        Dim sibling As String
        sibling = fso.GetParentFolderName(pythonExe) & "\pythonw.exe"
        If fso.FileExists(sibling) Then PreferPythonW = sibling
    End If
End Function

Private Function VbsQuoteString(ByVal value As String) As String
    VbsQuoteString = """" & Replace(value, """", """""") & """"
End Function

Private Function EscapePsPath(ByVal pathValue As String) As String
    EscapePsPath = Replace(pathValue, "'", "''")
End Function

Private Function RunHiddenCommand(ByVal command As String) As Long
    On Error GoTo EH

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    RunHiddenCommand = shell.Run(command, 0, True)
    Exit Function

EH:
    MSCANModLogging.WriteLog "RunHiddenCommand error: #" & Err.Number & " - " & Err.Description
    RunHiddenCommand = -1
End Function

' Legacy helper kept for PowerShell UTF-8 fallbacks.
Private Function RunCommandAndCaptureOutput(ByVal command As String, ByRef output As String, ByRef exitCode As Long, Optional ByVal timeoutMs As Long = 30000) As Boolean
    On Error GoTo EH
    RunCommandAndCaptureOutput = False ' Default to failure
    output = ""
    exitCode = -1

    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim process As Object

    ' Use Exec where available. Surround in On Error to capture failures.
    Set process = shell.Exec(command)

    Dim startTime As Double
    startTime = Timer

    ' Wait for the process to finish or timeout
    Do While process.status = 0
        DoEvents
        If timeoutMs > 0 Then
            If ((Timer - startTime) * 1000) > timeoutMs Then
                ' Timeout: attempt to terminate by using taskkill (best-effort)
                On Error Resume Next
                MSCANModLogging.WriteLog "RunCommandAndCaptureOutput: Process timed out after " & timeoutMs & "ms. Attempting to terminate."
                ' We can't reliably get PID from Exec, so best-effort fallback: taskkill by window title (not ideal).
                ' Log and return timeout.
                output = "Timeout waiting for external command."
                exitCode = -1
                RunCommandAndCaptureOutput = False
                Exit Function
            End If
        End If
    Loop

    exitCode = process.exitCode

    If exitCode = 0 Then
        If Not process.StdOut.AtEndOfStream Then
            output = process.StdOut.ReadAll()
        End If
    Else
        If Not process.StdErr.AtEndOfStream Then
            output = process.StdErr.ReadAll()
        Else
            ' If stderr empty, try stdout as a last resort
            If Not process.StdOut.AtEndOfStream Then output = process.StdOut.ReadAll()
        End If
    End If

    RunCommandAndCaptureOutput = True
    Exit Function

EH:
    output = "VBA Error in RunCommandAndCaptureOutput: #" & Err.Number & " - " & Err.Description
    RunCommandAndCaptureOutput = False
End Function


'===============================================================================
' HELPER: Inserts the generated HTML footer into the email body.
' Adds a duplication guard using a marker in the footer HTML.
'===============================================================================
' Routes a completed scan's footer into the mail. Returns True when the mail
' is in a final good state (applied or legitimately skipped); False means the
' apply failed and a retry with a freshly resolved item is worthwhile.
Private Function ApplyFooterToMail(ByVal mail As Object, ByVal footerPath As String, ByVal mode As String) As Boolean
    If mode = "full" Then
        If MailAlreadyHasFooter(mail) Then
            ApplyFooterToMail = ReplaceAesFooterInMail(mail, footerPath)
        Else
            ApplyFooterToMail = InsertFooterIntoMail(mail, footerPath)
        End If
    Else
        If MailAlreadyHasFooter(mail) Then
            MSCANModLogging.WriteLog "ApplyFooterToMail: compact footer already present, skipping: " & SafeSubject(mail)
            ApplyFooterToMail = True
        Else
            ApplyFooterToMail = InsertFooterIntoMail(mail, footerPath)
        End If
    End If
End Function

Private Function StripAesTopBannersFromBody(ByVal bodyHtml As String) As String
    ' Clear any previous status strip so re-scanning cannot stack duplicates.
    On Error Resume Next
    StripAesTopBannersFromBody = bodyHtml
    Const BANNER_START As String = "<!-- AES Top Banners Start -->"
    Const BANNER_END As String = "<!-- AES Top Banners End -->"
    Dim s As Long, e As Long
    Do
        s = InStr(1, bodyHtml, BANNER_START, vbTextCompare)
        If s = 0 Then Exit Do
        e = InStr(s, bodyHtml, BANNER_END, vbTextCompare)
        If e = 0 Then Exit Do
        e = e + Len(BANNER_END)
        bodyHtml = Left$(bodyHtml, s - 1) & Mid$(bodyHtml, e)
    Loop
    StripAesTopBannersFromBody = bodyHtml
End Function

' Removes the status strip from the footer payload and returns it. The strip
' belongs at the top of the body; the scan summary stays at the bottom.
Private Function TakeAesTopBannerBlock(ByRef footerHtml As String) As String
    On Error Resume Next
    Const BANNER_START As String = "<!-- AES Top Banners Start -->"
    Const BANNER_END As String = "<!-- AES Top Banners End -->"

    Dim s As Long, e As Long
    s = InStr(1, footerHtml, BANNER_START, vbTextCompare)
    If s = 0 Then Exit Function
    e = InStr(s, footerHtml, BANNER_END, vbTextCompare)
    If e = 0 Then Exit Function
    e = e + Len(BANNER_END)

    TakeAesTopBannerBlock = Mid$(footerHtml, s, e - s)
    footerHtml = Left$(footerHtml, s - 1) & Mid$(footerHtml, e)
End Function

Private Function AesBannerImagePath(ByVal bannerHtml As String) As String
    On Error Resume Next
    Dim s As Long, e As Long
    s = InStr(1, bannerHtml, AES_BANNER_MARKER, vbTextCompare)
    If s = 0 Then Exit Function
    s = s + Len(AES_BANNER_MARKER)
    e = InStr(s, bannerHtml, "-->")
    If e = 0 Then Exit Function
    AesBannerImagePath = Trim$(Mid$(bannerHtml, s, e - s))
End Function

Private Sub RemoveAesBannerAttachments(ByVal mail As Object)
    On Error Resume Next
    Dim i As Long
    For i = mail.Attachments.Count To 1 Step -1
        If InStr(1, mail.Attachments.Item(i).FileName, AES_BANNER_FILE_PREFIX, vbTextCompare) = 1 Then
            mail.Attachments.Item(i).Delete
        End If
    Next i
End Sub

' Attaches the rendered strip as a hidden inline image and repoints the markup
' at it. Falls back to the original file:// markup if the store rejects it.
Private Function EmbedAesBannerImage(ByVal mail As Object, ByVal bannerHtml As String) As String
    Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    Const PR_ATTACH_MIME_TAG As String = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
    Const PR_ATTACH_FLAGS As String = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
    Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
    Const ATT_MHTML_REF As Long = 4

    On Error GoTo ErrHandler
    EmbedAesBannerImage = bannerHtml

    Dim pngPath As String
    pngPath = AesBannerImagePath(bannerHtml)
    If Len(pngPath) = 0 Then Exit Function

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(pngPath) Then
        MSCANModLogging.WriteLog "EmbedAesBannerImage: banner image not found: " & pngPath
        Exit Function
    End If

    RemoveAesBannerAttachments mail

    Dim att As Object
    ' Position 0 keeps the part out of the body's attachment flow.
    Set att = mail.Attachments.Add(pngPath, olByValue, 0, fso.GetFileName(pngPath))
    mail.Save

    Dim pa As Object: Set pa = att.PropertyAccessor

    ' Without a Content-ID the cid: reference cannot resolve, so a rejection
    ' here must drop the attachment and leave the file:// markup alone.
    On Error Resume Next
    Err.Clear
    pa.SetProperty PR_ATTACH_CONTENT_ID, AES_BANNER_CID
    If Err.Number <> 0 Then
        MSCANModLogging.WriteLog "EmbedAesBannerImage: Content-ID rejected #" & Err.Number & " - " & Err.Description
        Err.Clear
        att.Delete
        mail.Save
        On Error GoTo ErrHandler
        Exit Function
    End If
    ' Presentation flags only — the image still resolves if these are refused.
    pa.SetProperty PR_ATTACH_MIME_TAG, "image/png"
    pa.SetProperty PR_ATTACH_FLAGS, ATT_MHTML_REF
    pa.SetProperty PR_ATTACHMENT_HIDDEN, True
    Err.Clear
    On Error GoTo ErrHandler
    mail.Save

    Dim fileUrl As String
    fileUrl = "file:///" & Replace(pngPath, "\", "/")
    EmbedAesBannerImage = Replace(bannerHtml, fileUrl, "cid:" & AES_BANNER_CID, 1, -1, vbTextCompare)
    Exit Function

ErrHandler:
    MSCANModLogging.WriteLog "EmbedAesBannerImage error: #" & Err.Number & " - " & Err.Description
    EmbedAesBannerImage = bannerHtml
End Function

Private Function InjectAesTopBanner(ByVal bodyHtml As String, ByVal bannerHtml As String) As String
    On Error Resume Next
    InjectAesTopBanner = bodyHtml
    If Len(bannerHtml) = 0 Then Exit Function

    Dim s As Long, e As Long
    s = InStr(1, bodyHtml, "<body", vbTextCompare)
    If s > 0 Then
        e = InStr(s, bodyHtml, ">")
        If e > 0 Then
            InjectAesTopBanner = Left$(bodyHtml, e) & vbCrLf & bannerHtml & Mid$(bodyHtml, e + 1)
            Exit Function
        End If
    End If

    InjectAesTopBanner = bannerHtml & bodyHtml
End Function

Private Function InsertFooterIntoMail(mail As Object, footerPath As String) As Boolean
    On Error GoTo ErrHandler
    InsertFooterIntoMail = False

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(footerPath) Then
        MSCANModLogging.WriteLog "InsertFooterIntoMail: Footer file not found: " & footerPath
        InsertFooterIntoMail = True   ' Nothing to apply; retrying will not help.
        Exit Function
    End If

    Dim footerHTML As String: footerHTML = ReadTextUtf8(footerPath)
    If Len(footerHTML) = 0 Then
        MSCANModLogging.WriteLog "InsertFooterIntoMail: Footer file is empty or could not be read: " & footerPath
        InsertFooterIntoMail = True
        Exit Function
    End If

    footerHTML = InjectFullScanAction(mail, footerHTML)

    Dim bannerHtml As String
    bannerHtml = TakeAesTopBannerBlock(footerHTML)

    ' Marker for duplication detection (AES; legacy GeoFooter still recognised)
    Const FOOTER_MARKER_AES As String = "<!-- AES Start -->"
    Const FOOTER_MARKER_LEGACY As String = "<!-- GeoFooter Start -->"

    If InStr(1, footerHTML, FOOTER_MARKER_AES, vbTextCompare) = 0 And _
       InStr(1, footerHTML, FOOTER_MARKER_LEGACY, vbTextCompare) = 0 Then
        footerHTML = FOOTER_MARKER_AES & vbCrLf & footerHTML
    End If

    If MailAlreadyHasFooter(mail) Then
        MarkAesScanned mail
        MSCANModLogging.WriteLog "InsertFooterIntoMail: Footer marker already present - skipping duplicate for subject: " & SafeSubject(mail)
        On Error Resume Next
        fso.DeleteFile footerPath
        On Error GoTo ErrHandler
        InsertFooterIntoMail = True
        Exit Function
    End If

    ' Normal mail: HTML body. ReportItem (ReadNotify IPNRN etc.): Body only.
    If TypeOf mail Is Outlook.MailItem Then
        ' Never force BodyFormat while an HTML body exists — on IMAP/Google
        ' stores that regenerates the body from the plain-text copy and
        ' flattens the whole message to text.
        Dim bodyHtml As String: bodyHtml = mail.HTMLBody
        If Len(Trim$(bodyHtml)) = 0 Then
            If mail.BodyFormat <> olFormatHTML Then mail.BodyFormat = olFormatHTML
            bodyHtml = mail.HTMLBody
        End If
        bodyHtml = StripAesTopBannersFromBody(bodyHtml)

        Dim newBody As String
        Dim pos As Long
        pos = InStrRev(bodyHtml, "</body>", , vbTextCompare)

        If pos > 0 Then
            newBody = Left$(bodyHtml, pos - 1) & footerHTML & Mid$(bodyHtml, pos)
        Else
            newBody = bodyHtml & "<hr>" & footerHTML
        End If

        ' Status strip goes at the top of the body, as an embedded image so the
        ' text never reaches the message-list preview column.
        If Len(bannerHtml) > 0 Then bannerHtml = EmbedAesBannerImage(mail, bannerHtml)
        mail.HTMLBody = InjectAesTopBanner(newBody, bannerHtml)
    Else
        Dim bodyText As String
        On Error Resume Next
        bodyText = CStr(mail.Body)
        Err.Clear
        On Error GoTo ErrHandler
        mail.Body = bodyText & vbCrLf & vbCrLf & footerHTML
    End If

    mail.Save
    MarkAesScanned mail
    MSCANModLogging.WriteLog "Footer inserted successfully into email with subject: " & SafeSubject(mail)

    On Error Resume Next
    fso.DeleteFile footerPath
    On Error GoTo ErrHandler

    InsertFooterIntoMail = True
    Exit Function

ErrHandler:
    MSCANModLogging.WriteLog "InsertFooterIntoMail error: #" & Err.Number & " - " & Err.Description
End Function

Private Function ReplaceAesFooterInMail(ByVal mail As Object, ByVal footerPath As String) As Boolean
    On Error GoTo ErrHandler
    ReplaceAesFooterInMail = False

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(footerPath) Then
        MSCANModLogging.WriteLog "ReplaceAesFooterInMail: Footer file not found: " & footerPath
        ReplaceAesFooterInMail = True   ' Nothing to apply; retrying will not help.
        Exit Function
    End If

    Dim footerHTML As String: footerHTML = ReadTextUtf8(footerPath)
    If Len(footerHTML) = 0 Then
        ReplaceAesFooterInMail = True
        Exit Function
    End If

    footerHTML = InjectFullScanAction(mail, footerHTML)

    Dim bannerHtml As String
    bannerHtml = TakeAesTopBannerBlock(footerHTML)

    Const FOOTER_START_AES As String = "<!-- AES Start -->"
    Const FOOTER_END_AES As String = "<!-- AES End -->"
    Const FOOTER_START_LEGACY As String = "<!-- GeoFooter Start -->"
    Const FOOTER_END_LEGACY As String = "<!-- GeoFooter End -->"

    ' ReportItem has no HTMLBody / BodyFormat — append/replace in Body text.
    If Not TypeOf mail Is Outlook.MailItem Then
        Dim rptBody As String
        On Error Resume Next
        rptBody = CStr(mail.Body)
        Err.Clear
        On Error GoTo ErrHandler
        Dim rptStart As Long, rptEnd As Long, rptEndMarker As String, rptEndLen As Long
        rptStart = InStr(1, rptBody, FOOTER_START_AES, vbTextCompare)
        rptEndMarker = FOOTER_END_AES
        If rptStart = 0 Then
            rptStart = InStr(1, rptBody, FOOTER_START_LEGACY, vbTextCompare)
            rptEndMarker = FOOTER_END_LEGACY
        End If
        If rptStart > 0 Then
            rptEnd = InStr(rptStart, rptBody, rptEndMarker, vbTextCompare)
            If rptEnd > 0 Then
                rptEndLen = Len(rptEndMarker)
                mail.Body = Left$(rptBody, rptStart - 1) & footerHTML & Mid$(rptBody, rptEnd + rptEndLen)
            Else
                mail.Body = Left$(rptBody, rptStart - 1) & footerHTML
            End If
        Else
            mail.Body = rptBody & vbCrLf & vbCrLf & footerHTML
        End If
        mail.Save
        MSCANModLogging.WriteLog "ReplaceAesFooterInMail: Full footer applied (report) for subject: " & SafeSubject(mail)
        On Error Resume Next
        fso.DeleteFile footerPath
        ReplaceAesFooterInMail = True
        Exit Function
    End If

    ' Never force BodyFormat while an HTML body exists — on IMAP/Google
    ' stores that regenerates the body from the plain-text copy and
    ' flattens the whole message to text.
    Dim bodyHtml As String: bodyHtml = mail.HTMLBody
    If Len(Trim$(bodyHtml)) = 0 Then
        If mail.BodyFormat <> olFormatHTML Then mail.BodyFormat = olFormatHTML
        bodyHtml = mail.HTMLBody
    End If
    bodyHtml = StripAesTopBannersFromBody(bodyHtml)
    Dim startPos As Long
    Dim endPos As Long
    Dim endMarker As String
    Dim endMarkerLen As Long

    startPos = InStr(1, bodyHtml, FOOTER_START_AES, vbTextCompare)
    endMarker = FOOTER_END_AES
    If startPos = 0 Then
        startPos = InStr(1, bodyHtml, FOOTER_START_LEGACY, vbTextCompare)
        endMarker = FOOTER_END_LEGACY
    End If

    If startPos = 0 Then
        ' Mail is marked scanned (AESScanned property) but the marker text is
        ' gone from the body — append the new footer instead of giving up.
        MSCANModLogging.WriteLog "ReplaceAesFooterInMail: No AES footer block found; appending instead."
        Dim insPos As Long
        Dim appendedBody As String
        insPos = InStrRev(bodyHtml, "</body>", , vbTextCompare)
        If insPos > 0 Then
            appendedBody = Left$(bodyHtml, insPos - 1) & footerHTML & Mid$(bodyHtml, insPos)
        Else
            appendedBody = bodyHtml & "<hr>" & footerHTML
        End If
        If Len(bannerHtml) > 0 Then bannerHtml = EmbedAesBannerImage(mail, bannerHtml)
        mail.HTMLBody = InjectAesTopBanner(appendedBody, bannerHtml)
        mail.Save
        MarkAesScanned mail
        On Error Resume Next
        fso.DeleteFile footerPath
        On Error GoTo ErrHandler
        ReplaceAesFooterInMail = True
        Exit Function
    End If

    endPos = InStr(startPos, bodyHtml, endMarker, vbTextCompare)
    If endPos = 0 Then
        MSCANModLogging.WriteLog "ReplaceAesFooterInMail: AES footer end marker not found."
        ReplaceAesFooterInMail = True   ' Structural, not transient — no retry.
        Exit Function
    End If

    endMarkerLen = Len(endMarker)
    Dim replacedBody As String
    replacedBody = Left$(bodyHtml, startPos - 1) & footerHTML & Mid$(bodyHtml, endPos + endMarkerLen)
    If Len(bannerHtml) > 0 Then bannerHtml = EmbedAesBannerImage(mail, bannerHtml)
    mail.HTMLBody = InjectAesTopBanner(replacedBody, bannerHtml)
    mail.Save

    MSCANModLogging.WriteLog "ReplaceAesFooterInMail: Full footer applied for subject: " & SafeSubject(mail)

    On Error Resume Next
    fso.DeleteFile footerPath
    On Error GoTo ErrHandler
    ReplaceAesFooterInMail = True
    Exit Function

ErrHandler:
    MSCANModLogging.WriteLog "ReplaceAesFooterInMail error: #" & Err.Number & " - " & Err.Description
End Function


'===============================================================================
' The following functions are mostly well-written and require minimal changes.
' They are included here for completeness.
'===============================================================================

Public Function GetInternetHeaders(mail As Object) As String
    On Error GoTo ErrHandler
    Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F" ' Fixed to Unicode
    GetInternetHeaders = mail.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    Exit Function
ErrHandler:
    MSCANModLogging.WriteLog "GetInternetHeaders error: #" & Err.Number & " - " & Err.Description
    GetInternetHeaders = ""
End Function

Private Sub WriteHeadersToFile(headers As String, filePath As String)
    On Error GoTo ErrHandler
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call EnsureFolderExists(fso.GetParentFolderName(filePath))
    WriteTextUtf8 filePath, headers
    Exit Sub
ErrHandler:
    MSCANModLogging.WriteLog "WriteHeadersToFile error: #" & Err.Number & " - " & Err.Description
End Sub

Private Function EnsureFolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    If Len(folderPath) = 0 Then Exit Function

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        EnsureFolderExists = True
        Exit Function
    End If

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And StrComp(parentPath, folderPath, vbTextCompare) <> 0 Then
        If Not EnsureFolderExists(parentPath) Then Exit Function
    End If

    fso.CreateFolder folderPath
    EnsureFolderExists = True
    Exit Function
EH:
    MSCANModLogging.WriteLog "EnsureFolderExists error creating '" & folderPath & "': #" & Err.Number & " - " & Err.Description
    EnsureFolderExists = False
End Function

Public Sub UpgradeToFullScanByEntryID(ByVal entryId As String)
    On Error GoTo EH

    If Len(entryId) = 0 Then
        MSCANModStatus.ShowStatus "AES Full Scan: missing email id."
        Exit Sub
    End If

    Dim mail As Object
    Set mail = Application.Session.GetItemFromID(entryId)
    If Not IsScannableItem(mail) Then
        MSCANModStatus.ShowStatus "AES Full Scan: could not open that email."
        Exit Sub
    End If

    MSCANModLogging.WriteLog "UpgradeToFullScanByEntryID: " & SafeSubject(mail)
    MSCANModStatus.ShowStatus "Running AES Full Scan: " & SafeSubject(mail)
    ProcessCompleteFooter mail

    If MailHasFullFooter(mail) Then
        MSCANModStatus.ShowStatus "AES Full Scan applied: " & SafeSubject(mail)
    Else
        MSCANModStatus.ShowStatus "AES Full Scan finished (check log): " & SafeSubject(mail)
    End If
    Exit Sub

EH:
    MSCANModLogging.WriteLog "UpgradeToFullScanByEntryID error: #" & Err.Number & " - " & Err.Description
    MSCANModStatus.ShowStatus "AES Full Scan failed: " & Err.Description
End Sub

Private Function InjectFullScanAction(ByVal mail As Object, ByVal footerHTML As String) As String
    On Error GoTo EH

    ' New footers already include a "Show Full Scan" link to a pre-built report.
    ' Only strip a leftover placeholder — do not launch a re-scan.
    Const MARKER As String = "{{AES_FULLSCAN}}"
    If InStr(1, footerHTML, MARKER, vbBinaryCompare) = 0 Then
        InjectFullScanAction = footerHTML
        Exit Function
    End If

    InjectFullScanAction = Replace(footerHTML, MARKER, _
        "<span style='text-decoration:underline;'>Show Full Scan</span>", 1, -1, vbBinaryCompare)
    MSCANModLogging.WriteLog "InjectFullScanAction: placeholder removed (report link expected from Python)."
    Exit Function

EH:
    MSCANModLogging.WriteLog "InjectFullScanAction error: #" & Err.Number & " - " & Err.Description
    InjectFullScanAction = Replace(footerHTML, "{{AES_FULLSCAN}}", _
        "<span style='text-decoration:underline;'>Show Full Scan</span>", 1, -1, vbBinaryCompare)
End Function

Private Function CreateFullScanLauncher(ByVal entryId As String) As String
    On Error GoTo EH

    Dim dir As String
    dir = GetBaseDir() & "\fullscan"
    If Not EnsureFolderExists(dir) Then
        CreateFullScanLauncher = ""
        Exit Function
    End If

    Dim fileName As String
    fileName = "fs_" & FullScanLauncherName(entryId) & ".vbs"

    Dim fullPath As String
    fullPath = dir & "\" & fileName

    Dim safeEntry As String
    safeEntry = Replace(entryId, """", """""")

    Dim script As String
    script = "On Error Resume Next" & vbCrLf & _
        "Dim olApp, entryId" & vbCrLf & _
        "entryId = """ & safeEntry & """" & vbCrLf & _
        "Set olApp = GetObject(, ""Outlook.Application"")" & vbCrLf & _
        "If olApp Is Nothing Then" & vbCrLf & _
        "  MsgBox ""Outlook must be running for AES Full Scan."", 48, ""AES""" & vbCrLf & _
        "  WScript.Quit 1" & vbCrLf & _
        "End If" & vbCrLf & _
        "olApp.AesUpgradeToFullScan entryId" & vbCrLf & _
        "If Err.Number <> 0 Then" & vbCrLf & _
        "  MsgBox ""AES Full Scan failed: "" & Err.Description, 48, ""AES""" & vbCrLf & _
        "End If" & vbCrLf

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(fullPath, True, False)
    ts.Write script
    ts.Close

    CreateFullScanLauncher = fullPath
    Exit Function

EH:
    MSCANModLogging.WriteLog "CreateFullScanLauncher error: #" & Err.Number & " - " & Err.Description
    CreateFullScanLauncher = ""
End Function

Private Function FullScanLauncherName(ByVal entryId As String) As String
    Dim cleaned As String
    Dim i As Long
    Dim ch As String

    cleaned = ""
    For i = 1 To Len(entryId)
        ch = Mid$(entryId, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            cleaned = cleaned & ch
        End If
        If Len(cleaned) >= 40 Then Exit For
    Next i

    If Len(cleaned) = 0 Then cleaned = Format$(Now, "yyyymmddhhnnss")
    FullScanLauncherName = cleaned
End Function

Public Function MailHasFooter(ByVal mail As Object) As Boolean
    MailHasFooter = MailAlreadyHasFooter(mail)
End Function

' Fast path for queue/catch-up — avoids reading HTMLBody.
Public Function IsAesScanned(ByVal mail As Object) As Boolean
    On Error Resume Next
    IsAesScanned = False
    If mail Is Nothing Then Exit Function
    Dim up As Object
    Set up = mail.UserProperties.Find(AES_SCANNED_PROP)
    If Not up Is Nothing Then
        IsAesScanned = (CStr(up.Value) = "1")
    End If
End Function

Public Sub MarkAesScanned(ByVal mail As Object)
    On Error Resume Next
    If mail Is Nothing Then Exit Sub
    Dim ups As Object
    Dim up As Object
    Set ups = mail.UserProperties
    If ups Is Nothing Then Exit Sub
    Set up = ups.Find(AES_SCANNED_PROP)
    If up Is Nothing Then Set up = ups.Add(AES_SCANNED_PROP, 1) ' olText
    If Not up Is Nothing Then
        up.Value = "1"
        mail.Save
    End If
End Sub

Public Function MailHasFullFooter(ByVal mail As Object) As Boolean
    On Error Resume Next
    Const FULL_MARKER As String = "<!-- AES Full Start -->"
    Dim bodyText As String
    bodyText = ItemBodyText(mail)
    MailHasFullFooter = (InStr(1, bodyText, FULL_MARKER, vbTextCompare) > 0)
End Function

Private Function MailAlreadyHasFooter(ByVal mail As Object) As Boolean
    On Error Resume Next
    MailAlreadyHasFooter = False
    If IsAesScanned(mail) Then
        MailAlreadyHasFooter = True
        Exit Function
    End If
    Const FOOTER_MARKER_AES As String = "<!-- AES Start -->"
    Const FOOTER_MARKER_LEGACY As String = "<!-- GeoFooter Start -->"
    Dim bodyText As String
    bodyText = ItemBodyText(mail)
    MailAlreadyHasFooter = (InStr(1, bodyText, FOOTER_MARKER_AES, vbTextCompare) > 0 Or _
                            InStr(1, bodyText, FOOTER_MARKER_LEGACY, vbTextCompare) > 0)
End Function

Private Function ItemBodyText(ByVal mail As Object) As String
    On Error Resume Next
    ItemBodyText = ""
    If mail Is Nothing Then Exit Function
    If TypeOf mail Is Outlook.MailItem Then
        If mail.BodyFormat = olFormatHTML Then
            ItemBodyText = mail.HTMLBody
        Else
            ItemBodyText = mail.Body
        End If
    Else
        ItemBodyText = CStr(mail.Body)
    End If
End Function

Private Function SafeSubject(mail As Object) As String
    On Error Resume Next
    SafeSubject = mail.Subject
    If Err.Number <> 0 Or Len(SafeSubject) = 0 Then SafeSubject = "[No Subject]"
End Function

' --- UTF-8 Read/Write Utilities (Unchanged, as they are already robust) ---

Private Sub WriteTextUtf8(ByVal filePath As String, ByVal content As String)
    On Error GoTo EH
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    stm.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stm.Close
    Exit Sub
EH:
    MSCANModLogging.WriteLog "WriteTextUtf8: ADODB.Stream failed. Error #" & Err.Number & " - " & Err.Description & ". Falling back to PowerShell."
    WriteTextUtf8_PS filePath, content
End Sub

Private Function ReadTextUtf8(ByVal filePath As String) As String
    On Error GoTo EH
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextUtf8 = stm.ReadText(-1) ' adReadAll
    stm.Close
    Exit Function
EH:
    MSCANModLogging.WriteLog "ReadTextUtf8: ADODB.Stream failed. Error #" & Err.Number & " - " & Err.Description & ". Falling back to PowerShell."
    ReadTextUtf8 = ReadTextUtf8_PS(filePath)
End Function

Private Sub WriteTextUtf8_PS(ByVal filePath As String, ByVal content As String)
    Dim psCmd As String
    psCmd = "$content = '" & Replace(content, "'", "''") & "'; $content | Out-File -LiteralPath '" & filePath & "' -Encoding utf8 -NoNewline"
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & psCmd, 0, True
End Sub

Private Function ReadTextUtf8_PS(ByVal filePath As String) As String
    Dim output As String, exitCode As Long
    Dim psCmd As String: psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""Get-Content -LiteralPath '" & filePath & "' -Raw"""
    If RunCommandAndCaptureOutput(psCmd, output, exitCode) And exitCode = 0 Then
        ReadTextUtf8_PS = output
    Else
        MSCANModLogging.WriteLog "ReadTextUtf8_PS: PowerShell fallback failed to read file: " & filePath
        ReadTextUtf8_PS = ""
    End If
End Function

